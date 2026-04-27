using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

namespace HMVTools
{
    // ══════════════════════════════════════════════════════════════════════════
    //  ENTRY POINT
    // ══════════════════════════════════════════════════════════════════════════
    [Transaction(TransactionMode.Manual)]
    public class ReshapeFlexPipeCommand : IExternalCommand
    {
        private static ReshapeFlexPipeHandler _handler;
        private static ExternalEvent          _exEvent;
        private static ReshapeFlexPipeWindow  _window;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            View activeView = commandData.Application.ActiveUIDocument.ActiveView;
            if (activeView.ViewType != ViewType.Section && activeView.ViewType != ViewType.Detail)
            {
                TaskDialog.Show("HMV Tools – Reshape FlexPipe",
                    "This command must be launched from a Section or Detail view.\n\n" +
                    "Open a Section / Detail view and try again.");
                return Result.Cancelled;
            }

            if (_handler == null)
            {
                _handler = new ReshapeFlexPipeHandler();
                _exEvent  = ExternalEvent.Create(_handler);
            }

            if (_window != null && _window.IsLoaded)
            {
                _window.Focus();
                return Result.Succeeded;
            }

            _window = new ReshapeFlexPipeWindow(commandData.Application, _handler, _exEvent);

            var helper = new System.Windows.Interop.WindowInteropHelper(_window);
            helper.Owner = commandData.Application.MainWindowHandle;

            _window.Show();
            return Result.Succeeded;
        }

        public static void ClearWindow() => _window = null;
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  SELECTION FILTERS
    // ══════════════════════════════════════════════════════════════════════════

    public class ReshapeFlexPipeFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)        => elem is FlexPipe;
        public bool AllowReference(Reference r, XYZ p) => false;
    }

    public class ReshapeCurveFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)        => elem is CurveElement;
        public bool AllowReference(Reference r, XYZ p) => false;
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  PICK HANDLER
    // ══════════════════════════════════════════════════════════════════════════
    public class ReshapePickHandler : IExternalEventHandler
    {
        public enum PickStep { FlexPipe, Curves }

        public PickStep              Step { get; set; }
        public ReshapeFlexPipeWindow UI   { get; set; }

        public ElementId        PickedFlexPipeId  { get; private set; }
        public List<ElementId>  PickedCurveIds    { get; private set; } = new List<ElementId>();

        public string GetName() => "ReshapePickHandler";

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document   doc   = uidoc.Document;

            try
            {
                if (Step == PickStep.FlexPipe)
                {
                    Reference r = uidoc.Selection.PickObject(
                        ObjectType.Element,
                        new ReshapeFlexPipeFilter(),
                        "Select the FlexPipe to reshape — ESC to cancel");

                    PickedFlexPipeId = r.ElementId;
                    UI?.OnFlexPipePicked(doc.GetElement(r.ElementId) as FlexPipe);
                }
                else
                {
                    PickedCurveIds.Clear();

                    IList<Reference> refs = uidoc.Selection.PickObjects(
                        ObjectType.Element,
                        new ReshapeCurveFilter(),
                        "Select one or more Lines / Curves for the new path — finish with ✔");

                    foreach (Reference r in refs)
                        PickedCurveIds.Add(r.ElementId);

                    UI?.OnCurvesPicked(PickedCurveIds.Count);
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                UI?.SetStatus("Selection cancelled.");
                UI?.RestoreWindow();
            }
            catch (Exception ex)
            {
                UI?.SetStatus("Pick error: " + ex.Message);
                UI?.RestoreWindow();
            }
        }
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  MAIN RESHAPE HANDLER
    // ══════════════════════════════════════════════════════════════════════════
    public class ReshapeFlexPipeHandler : IExternalEventHandler
    {
        // ── Set by the UI before Raise() ──────────────────────────────────────
        public ReshapeFlexPipeWindow UI               { get; set; }
        public ElementId             FlexPipeId       { get; set; }
        public List<ElementId>       CurveIds         { get; set; }
        public double                ResampleInterval { get; set; } = 0.5; // feet

        public string GetName() => "ReshapeFlexPipeHandler";

        // ── Internal parameter snapshot struct ───────────────────────────────
        private struct ParamSnapshot
        {
            public string      Name;
            public StorageType Storage;
            public string      StrVal;
            public double      DblVal;
            public int         IntVal;
        }

        // ════════════════════════════════════════════════════════════════════
        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;

            if (FlexPipeId == null || CurveIds == null || CurveIds.Count == 0)
            {
                UI?.SetStatus("Error: Select both a FlexPipe and at least one curve.");
                return;
            }

            try
            {
                UI?.SetStatus("Reading existing FlexPipe...");

                FlexPipe oldPipe = doc.GetElement(FlexPipeId) as FlexPipe;
                if (oldPipe == null)
                    throw new InvalidOperationException("The selected FlexPipe no longer exists.");

                // ── 1. Clone all data from the old pipe ───────────────────────
                ElementId           typeId    = oldPipe.FlexPipeType.Id;
                ElementId           levelId   = GetLevelId(doc, oldPipe);
                ElementId           systemId  = GetSystemTypeId(doc, oldPipe);
                IList<XYZ>          oldPts    = oldPipe.Points;
                List<ParamSnapshot> snapshots = SnapshotParameters(oldPipe);

                if (oldPts == null || oldPts.Count < 2)
                    throw new InvalidOperationException("The existing FlexPipe has no valid points.");

                // ptA / ptB = first and last point of the original pipe
                // These define the horizontal plane of the FlexPipe.
                // NOTHING moves in X/Y — we only reshape the Z profile.
                XYZ ptA = oldPts[0];
                XYZ ptB = oldPts[oldPts.Count - 1];

                // ── 2. Define the pipe's vertical curtain plane ───────────────
                // uAB  = horizontal direction of the pipe (floor-plan axis)
                // BasisZ = vertical axis
                // The curtain is the plane that contains both uAB and Z through ptA.
                XYZ abVec = ptB - ptA;
                double abLen = abVec.GetLength();
                if (abLen < 1e-6)
                    throw new InvalidOperationException("FlexPipe start/end points are coincident.");

                XYZ uAB = abVec.Normalize();

                // ── 3. Extract raw points from the selected section curves ─────
                UI?.SetStatus("Extracting curve points...");
                List<XYZ> rawCurve = ExtractCurvePoints(doc, CurveIds);

                if (rawCurve.Count < 2)
                    throw new InvalidOperationException(
                        "Could not extract at least 2 points from the selected curves.");

                // ── 4. Project curve points onto the curtain plane ────────────
                // For each raw curve point P (picked in the section view):
                //   • "along"  = how far along uAB from ptA  → stays as horizontal displacement
                //   • "height" = the Z coordinate of P       → becomes the new Z of the pipe point
                //
                // The section curve lives in a plane that may have any origin.
                // We care only about the relative "along" and "Z" values,
                // scaled so that the first point maps to ptA and the last to ptB.

                List<XYZ> projected = ProjectOntoCurtain(rawCurve, ptA, ptB, uAB);

                // ── 5. Resample ───────────────────────────────────────────────
                UI?.SetStatus("Resampling trajectory...");
                List<XYZ> trajectory = ResampleTrajectory(projected, ResampleInterval);

                // Enforce clear min-distance from endpoints
                double minClearance = ResampleInterval;
                while (trajectory.Count > 2 && trajectory[0].DistanceTo(ptA) < minClearance)
                    trajectory.RemoveAt(0);
                while (trajectory.Count > 2 && trajectory[trajectory.Count - 1].DistanceTo(ptB) < minClearance)
                    trajectory.RemoveAt(trajectory.Count - 1);

                // Always start/end at the exact original A/B points
                trajectory.Insert(0, ptA);
                trajectory.Add(ptB);

                if (trajectory.Count < 2)
                    throw new InvalidOperationException("Trajectory is too short after resampling.");

                // ── 6. Tangents ───────────────────────────────────────────────
                XYZ startTangent = (trajectory[1]                          - trajectory[0]).Normalize();
                XYZ endTangent   = (trajectory[trajectory.Count - 1] - trajectory[trajectory.Count - 2]).Normalize();

                // ── 7. Transaction: delete old, create new ────────────────────
                UI?.SetStatus("Rebuilding FlexPipe...");
                using (Transaction t = new Transaction(doc, "Reshape FlexPipe"))
                {
                    t.Start();

                    FailureHandlingOptions fho = t.GetFailureHandlingOptions();
                    fho.SetFailuresPreprocessor(new WarningSuppressor());
                    t.SetFailureHandlingOptions(fho);

                    // Delete old pipe FIRST so the new one can reuse connectors
                    doc.Delete(FlexPipeId);

                    FlexPipe newPipe = FlexPipe.Create(
                        doc,
                        systemId,
                        typeId,
                        levelId,
                        startTangent,
                        endTangent,
                        trajectory);

                    if (newPipe == null)
                        throw new InvalidOperationException("FlexPipe.Create returned null. Check system/type IDs.");

                    // Restore all snapshotted parameters
                    RestoreParameters(newPipe, snapshots);

                    t.Commit();
                }

                UI?.SetStatus($"✔ FlexPipe reshaped successfully. ({trajectory.Count} points)");
            }
            catch (Exception ex)
            {
                UI?.SetStatus("Error: " + ex.Message);
            }
        }

        // ════════════════════════════════════════════════════════════════════
        //  CURVE POINT EXTRACTION
        //  Orders multiple selected CurveElements into a single polyline by
        //  chaining them end-to-end (nearest endpoint matching).
        // ════════════════════════════════════════════════════════════════════
        private static List<XYZ> ExtractCurvePoints(Document doc, List<ElementId> ids)
        {
            // Tessellate each curve into a list of XYZ
            var segments = new List<List<XYZ>>();
            foreach (ElementId id in ids)
            {
                CurveElement ce = doc.GetElement(id) as CurveElement;
                if (ce?.GeometryCurve == null) continue;

                IList<XYZ> tess = ce.GeometryCurve.Tessellate();
                if (tess != null && tess.Count >= 2)
                    segments.Add(tess.ToList());
            }

            if (segments.Count == 0) return new List<XYZ>();
            if (segments.Count == 1) return CleanPoints(segments[0]);

            // Chain segments: always append to a growing polyline by matching
            // the nearest endpoint of each unused segment to the current tail.
            var chain = new List<XYZ>(segments[0]);
            var remaining = segments.Skip(1).ToList();

            while (remaining.Count > 0)
            {
                XYZ tail = chain[chain.Count - 1];
                int bestIdx  = -1;
                bool reverse = false;
                double bestDist = double.MaxValue;

                for (int i = 0; i < remaining.Count; i++)
                {
                    double dStart = tail.DistanceTo(remaining[i][0]);
                    double dEnd   = tail.DistanceTo(remaining[i][remaining[i].Count - 1]);

                    if (dStart < bestDist) { bestDist = dStart; bestIdx = i; reverse = false; }
                    if (dEnd   < bestDist) { bestDist = dEnd;   bestIdx = i; reverse = true;  }
                }

                List<XYZ> seg = remaining[bestIdx];
                if (reverse) seg = Enumerable.Reverse(seg).ToList();
                remaining.RemoveAt(bestIdx);

                // Skip first point to avoid exact duplicate
                chain.AddRange(seg.Skip(1));
            }

            return CleanPoints(chain);
        }

        // ════════════════════════════════════════════════════════════════════
        //  CURTAIN PROJECTION
        //
        //  The FlexPipe lies horizontally in the floor plan.
        //  uAB  = direction from ptA to ptB in plan (the pipe's horizontal axis).
        //  BasisZ = global up.
        //
        //  The section/detail curve represents the vertical profile of the pipe
        //  viewed from the side.  We interpret the curve in 2-D:
        //    • local U (horizontal in section) → maps to uAB in 3-D
        //    • local V (vertical in section)   → maps to BasisZ in 3-D
        //
        //  We normalise the curve's own U-span to the real A-B distance,
        //  keeping the V (Z) values as absolute displacements added to ptA.Z.
        // ════════════════════════════════════════════════════════════════════
        private static List<XYZ> ProjectOntoCurtain(
            List<XYZ> curvePts, XYZ ptA, XYZ ptB, XYZ uAB)
        {
            if (curvePts.Count < 2)
                throw new InvalidOperationException("Curve has fewer than 2 points.");

            // Find the "horizontal" axis of the section view by looking at the
            // maximum spread of the curve points.  We project every point onto
            // XY to find the dominant in-plane direction of the section curve.
            // Then we treat the curve's own start→end vector as "along" and
            // the Z component as "height".

            XYZ cStart = curvePts[0];
            XYZ cEnd   = curvePts[curvePts.Count - 1];

            // Section-local horizontal direction (projected onto XY)
            XYZ cHoriz = new XYZ(cEnd.X - cStart.X, cEnd.Y - cStart.Y, 0.0);
            double cHorizLen = cHoriz.GetLength();

            // If the curve is drawn exactly vertically (pure Z) in the section,
            // we cannot determine a horizontal span — treat it as a straight
            // vertical rise.  Use uAB for the along direction with zero span.
            bool pureVertical = cHorizLen < 1e-6;

            // Total real-world horizontal distance A→B
            double abLen = ptB.DistanceTo(ptA);   // ptA and ptB share Z (flat pipe)

            var result = new List<XYZ>();
            foreach (XYZ cp in curvePts)
            {
                double along, height;

                if (pureVertical)
                {
                    // Map Z of the curve linearly from 0 to abLen
                    double t = (curvePts.Count > 1)
                        ? (cp.Z - cStart.Z) / Math.Max(1e-9, cEnd.Z - cStart.Z)
                        : 0.0;
                    along  = t * abLen;
                    height = cp.Z;
                }
                else
                {
                    // Project cp onto the section horizontal direction
                    XYZ uHoriz = cHoriz.Normalize();
                    XYZ local  = cp - cStart;
                    along  = local.DotProduct(uHoriz) * (abLen / cHorizLen); // scaled to real A-B
                    height = cp.Z;                                             // absolute Z kept
                }

                // Place onto the curtain plane:
                // X/Y = ptA + uAB * along   (stays on the floor-plan line)
                // Z   = height (from the section curve)
                XYZ pt3D = new XYZ(
                    ptA.X + uAB.X * along,
                    ptA.Y + uAB.Y * along,
                    height);

                result.Add(pt3D);
            }

            // Force exact A and B endpoints
            result[0]                  = ptA;
            result[result.Count - 1]   = ptB;

            return result;
        }

        // ════════════════════════════════════════════════════════════════════
        //  RESAMPLE  (identical logic to ElectricalConnectionCommand)
        // ════════════════════════════════════════════════════════════════════
        private static List<XYZ> ResampleTrajectory(List<XYZ> points, double interval)
        {
            double totalLen = 0;
            for (int i = 1; i < points.Count; i++)
                totalLen += points[i].DistanceTo(points[i - 1]);

            if (totalLen < interval * 2) return new List<XYZ>(points);

            var result = new List<XYZ> { points[0] };
            double sinceLastSample = 0;

            for (int i = 1; i < points.Count; i++)
            {
                double segLen = points[i].DistanceTo(points[i - 1]);
                XYZ    dir    = (points[i] - points[i - 1]).Normalize();
                double consumed = 0;

                while (consumed < segLen)
                {
                    double remaining = interval - sinceLastSample;
                    double available = segLen - consumed;

                    if (available >= remaining)
                    {
                        result.Add(points[i - 1] + dir * (consumed + remaining));
                        consumed += remaining;
                        sinceLastSample = 0;
                    }
                    else
                    {
                        sinceLastSample += available;
                        consumed = segLen;
                    }
                }
            }

            if (result[result.Count - 1].DistanceTo(points[points.Count - 1]) > 0.001)
                result.Add(points[points.Count - 1]);

            return result;
        }

        // ════════════════════════════════════════════════════════════════════
        //  PARAMETER SNAPSHOT / RESTORE
        // ════════════════════════════════════════════════════════════════════
        private static List<ParamSnapshot> SnapshotParameters(Element elem)
        {
            var list = new List<ParamSnapshot>();
            foreach (Parameter p in elem.Parameters)
            {
                if (p.IsReadOnly || p.Definition == null) continue;
                if (string.IsNullOrWhiteSpace(p.Definition.Name)) continue;

                var snap = new ParamSnapshot
                {
                    Name    = p.Definition.Name,
                    Storage = p.StorageType
                };

                switch (p.StorageType)
                {
                    case StorageType.String:
                        snap.StrVal = p.AsString();
                        break;
                    case StorageType.Double:
                        snap.DblVal = p.AsDouble();
                        break;
                    case StorageType.Integer:
                        snap.IntVal = p.AsInteger();
                        break;
                    default:
                        continue; // skip ElementId params
                }
                list.Add(snap);
            }
            return list;
        }

        private static void RestoreParameters(Element elem, List<ParamSnapshot> snapshots)
        {
            foreach (ParamSnapshot snap in snapshots)
            {
                Parameter p = elem.LookupParameter(snap.Name);
                if (p == null || p.IsReadOnly) continue;

                try
                {
                    switch (snap.Storage)
                    {
                        case StorageType.String:
                            if (snap.StrVal != null) p.Set(snap.StrVal);
                            break;
                        case StorageType.Double:
                            p.Set(snap.DblVal);
                            break;
                        case StorageType.Integer:
                            p.Set(snap.IntVal);
                            break;
                    }
                }
                catch { /* skip parameters that reject the value (e.g. computed) */ }
            }
        }

        // ════════════════════════════════════════════════════════════════════
        //  HELPERS
        // ════════════════════════════════════════════════════════════════════
        private static ElementId GetLevelId(Document doc, FlexPipe pipe)
        {
            Level lvl = pipe.ReferenceLevel;
            if (lvl != null) return lvl.Id;

            return new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => l.Elevation)
                .FirstOrDefault()?.Id
                ?? ElementId.InvalidElementId;
        }

        private static ElementId GetSystemTypeId(Document doc, FlexPipe pipe)
        {
            // A FlexPipe always belongs to a piping system.
            // Read the system from its connector and get the SystemType.
            ConnectorManager cm = pipe.ConnectorManager;
            if (cm != null)
            {
                foreach (Connector c in cm.Connectors)
                {
                    if (c.MEPSystem is PipingSystem ps)
                        return ps.GetTypeId();
                }
            }

            // Fallback: first PipingSystemType in the document
            return new FilteredElementCollector(doc)
                .OfClass(typeof(PipingSystemType))
                .FirstElementId();
        }

        private static List<XYZ> CleanPoints(List<XYZ> pts)
        {
            if (pts.Count == 0) return pts;
            var cleaned = new List<XYZ> { pts[0] };
            for (int i = 1; i < pts.Count; i++)
                if (pts[i].DistanceTo(cleaned[cleaned.Count - 1]) > 0.001)
                    cleaned.Add(pts[i]);
            return cleaned;
        }
    }
}