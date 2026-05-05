using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using IxMilia.Dxf;
using IxMilia.Dxf.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HMVTools
{
    // ═══════════════════════════════════════════════════════════════════
    //  CONFIGURATION
    // ═══════════════════════════════════════════════════════════════════

    public class ElectricalRefresherConfig
    {
        // Parameters on adaptive point
        public string CnxNumberParam    { get; set; } = "HMV_CFI_CONEXI\u00d3N";
        public string EquipoInicialParam { get; set; } = "HMV_CFI_EQUIPO INICIAL";
        public string ConnectedParam    { get; set; } = "HMV_CFI_CONNECTED";

        // Parameters on FlexPipe
        public string FlexCnxNumberParam    { get; set; } = "HMV_CFI_CONEXI\u00d3N";
        public string FlexEquipoInicialParam { get; set; } = "HMV_CFI_EQUIPO INICIAL";
        public string FlexEquipoFinalParam   { get; set; } = "HMV_CFI_EQUIPO FINAL";

        // Types
        public string FlexPipeTypeKey        { get; set; }
        public string ElectricalSystemTypeKey { get; set; }

        // DXF folder
        public string DxfFolder { get; set; }
    }

    // ═══════════════════════════════════════════════════════════════════
    //  COMMAND  — ribbon entry point
    // ═══════════════════════════════════════════════════════════════════

    [Transaction(TransactionMode.Manual)]
    public class ElectricalRefresherCommand : IExternalCommand
    {
        private static ElectricalRefresherPickHandler      _pickHandler   = null;
        private static ExternalEvent                        _pickEvent     = null;
        private static ElectricalRefresherHandler           _runHandler    = null;
        private static ExternalEvent                        _runEvent      = null;
        private static MirrorFlexPipeRefresherHandler       _mirrorHandler = null;
        private static ExternalEvent                        _mirrorEvent   = null;
        private static ElectricalRefresherWindow            _window        = null;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (_window != null && _window.IsLoaded)
            {
                _window.Focus();
                return Result.Succeeded;
            }

            _pickHandler   = new ElectricalRefresherPickHandler();
            _pickEvent     = ExternalEvent.Create(_pickHandler);
            _runHandler    = new ElectricalRefresherHandler();
            _runEvent      = ExternalEvent.Create(_runHandler);
            _mirrorHandler = new MirrorFlexPipeRefresherHandler();
            _mirrorEvent   = ExternalEvent.Create(_mirrorHandler);

            _window = new ElectricalRefresherWindow(
                commandData.Application,
                _pickHandler,   _pickEvent,
                _runHandler,    _runEvent,
                _mirrorHandler, _mirrorEvent);

            var helper = new System.Windows.Interop.WindowInteropHelper(_window);
            helper.Owner = commandData.Application.MainWindowHandle;

            _window.Show();
            return Result.Succeeded;
        }

        public static void ClearWindow() { _window = null; }
    }

    // ═══════════════════════════════════════════════════════════════════
    //  PICK HANDLER
    // ═══════════════════════════════════════════════════════════════════

    public class ElectricalRefresherPickHandler : IExternalEventHandler
    {
        public ElectricalRefresherWindow UI { get; set; }

        public string GetName() => "ElectricalRefresherPickHandler";

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;

            try
            {
                IList<Reference> refs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    new AdaptiveComponentFilter(),
                    "Select Adaptive Points — click Finish when done");

                if (refs == null || refs.Count == 0)
                {
                    UI?.SetStatus("No elements selected.");
                    UI?.RestoreWindow();
                    return;
                }

                UI?.OnPointsPicked(refs.Select(r => r.ElementId).ToList());
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

    // ═══════════════════════════════════════════════════════════════════
    //  MAIN HANDLER
    // ═══════════════════════════════════════════════════════════════════

    public class ElectricalRefresherHandler : IExternalEventHandler
    {
        public ElectricalRefresherWindow UI     { get; set; }
        public List<ElementId>           PointIds { get; set; }
        public ElectricalRefresherConfig Config  { get; set; }
        public bool                      IsReset { get; set; }

        public string GetName() => "ElectricalRefresherHandler";

        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;

            if (PointIds == null || PointIds.Count == 0 || Config == null)
            {
                UI?.Log("Error: No points or configuration.");
                return;
            }

            if (IsReset)
                ExecuteReset(doc);
            else
                ExecuteRefresh(doc);
        }

        // ─────────────────────────────────────────────────────────────
        //  RESET MODE — delete FlexPipes whose CNX_NUMBER matches the
        //               selected adaptive points
        // ─────────────────────────────────────────────────────────────
        private void ExecuteReset(Document doc)
        {
            UI?.Log("── RESET MODE ──────────────────────────────────");

            // Collect connection numbers from selected points
            var cnxNumbers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ElementId id in PointIds)
            {
                FamilyInstance fi = doc.GetElement(id) as FamilyInstance;
                if (fi == null) continue;
                string cnx = GetStringParam(fi, Config.CnxNumberParam);
                if (!string.IsNullOrWhiteSpace(cnx)) cnxNumbers.Add(cnx);
            }

            if (cnxNumbers.Count == 0)
            {
                UI?.Log("Warning: No connection numbers found on selected points.");
                return;
            }

            // Collect FlexPipes matching those numbers
            var flexPipesToDelete = new FilteredElementCollector(doc)
                .OfClass(typeof(FlexPipe))
                .Cast<FlexPipe>()
                .Where(fp =>
                {
                    string v = GetStringParam(fp, Config.FlexCnxNumberParam);
                    return !string.IsNullOrWhiteSpace(v) && cnxNumbers.Contains(v);
                })
                .Select(fp => fp.Id)
                .ToList();

            using (Transaction trans = new Transaction(doc, "HMV - Reset Electrical Connections"))
            {
                trans.Start();

                FailureHandlingOptions fho = trans.GetFailureHandlingOptions();
                fho.SetFailuresPreprocessor(new WarningSuppressor());
                trans.SetFailureHandlingOptions(fho);

                int deleted = 0;
                foreach (ElementId fpId in flexPipesToDelete)
                {
                    try { doc.Delete(fpId); deleted++; }
                    catch (Exception ex) { UI?.Log($"  Could not delete FlexPipe {fpId}: {ex.Message}"); }
                }

                // Reset HMV_CFI_CONNECTED = 0 on the points
                int reset = 0;
                foreach (ElementId id in PointIds)
                {
                    FamilyInstance fi = doc.GetElement(id) as FamilyInstance;
                    if (fi == null) continue;
                    Parameter p = fi.LookupParameter(Config.ConnectedParam);
                    if (p != null && !p.IsReadOnly && p.StorageType == StorageType.Integer)
                    { p.Set(0); reset++; }
                }

                trans.Commit();
                UI?.Log($"✔ Deleted {deleted} FlexPipe(s). Reset flag on {reset} point(s).");
            }
        }

        // ─────────────────────────────────────────────────────────────
        //  REFRESH MODE — pair points by CNX_NUMBER and create FlexPipes
        // ─────────────────────────────────────────────────────────────
        private void ExecuteRefresh(Document doc)
        {
            UI?.Log("── REFRESH MODE ─────────────────────────────────");

            if (string.IsNullOrWhiteSpace(Config.DxfFolder) || !Directory.Exists(Config.DxfFolder))
            {
                UI?.Log("Error: DXF folder not found — check configuration.");
                return;
            }

            // ── Level ──────────────────────────────────────────────────────────
            Level defaultLevel = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => l.Elevation)
                .FirstOrDefault();

            if (defaultLevel == null) { UI?.Log("Error: No Level found in project."); return; }

            // ── Resolve FlexPipe / System types ────────────────────────────────
            FlexPipeType flexType = FindFlexPipeType(doc, Config.FlexPipeTypeKey);
            if (flexType == null) { UI?.Log($"Error: FlexPipe type '{Config.FlexPipeTypeKey}' not found."); return; }

            ElementId systemTypeId = FindElectricalSystemTypeId(doc, Config.ElectricalSystemTypeKey);
            if (systemTypeId == null) { UI?.Log($"Error: System type '{Config.ElectricalSystemTypeKey}' not found."); return; }

            // ── Build DXF map: CNX_NUMBER → file path ──────────────────────────
            Dictionary<string, string> dxfMap = BuildDxfMap(Config.DxfFolder);

            // ── Collect adaptive points with their CNX_NUMBER ──────────────────
            var pointData = new List<(ElementId id, FamilyInstance fi, string cnx, string equipoId, XYZ pos)>();

            foreach (ElementId id in PointIds)
            {
                FamilyInstance fi = doc.GetElement(id) as FamilyInstance;
                if (fi == null) continue;

                string cnx = GetStringParam(fi, Config.CnxNumberParam);
                if (string.IsNullOrWhiteSpace(cnx)) { UI?.Log($"  Skip {id}: no {Config.CnxNumberParam}."); continue; }

                // Only process points where IS CONNECTED = false (0)
                Parameter connParam = fi.LookupParameter(Config.ConnectedParam);
                if (connParam != null && connParam.StorageType == StorageType.Integer && connParam.AsInteger() == 1)
                {
                    UI?.Log($"  Skip {id} [{cnx}]: already connected.");
                    continue;
                }

                string equipoId = GetStringParam(fi, Config.EquipoInicialParam);
                XYZ pos = GetAdaptivePosition(doc, fi);
                pointData.Add((id, fi, cnx, equipoId, pos));
            }

            // ── Group by CNX_NUMBER ────────────────────────────────────────────
            var groups = pointData
                .GroupBy(p => p.cnx, StringComparer.OrdinalIgnoreCase)
                .ToList();

            int created = 0;
            int skipped = 0;

            foreach (var group in groups)
            {
                var pts = group.ToList();
                string cnx = group.Key;

                if (pts.Count < 2)
                {
                    UI?.Log($"  [{cnx}] Skipped: only {pts.Count} point(s) found.");
                    skipped++;
                    continue;
                }

                // Find optimal pairs — filters same-equipment pairs, minimises distance
                var pairIndices = FindOptimalPairIndices(pts);

                if (pairIndices.Count == 0)
                {
                    UI?.Log($"  [{cnx}] Skipped: {pts.Count} point(s) but no valid pairs (all share the same equipment).");
                    skipped++;
                    continue;
                }

                if (pts.Count > 2)
                    UI?.Log($"  [{cnx}] {pts.Count} points → {pairIndices.Count} pair(s) to connect.");

                // ── Find DXF file (shared by all pairs in this group) ──────────
                if (!dxfMap.TryGetValue(cnx, out string dxfPath))
                {
                    UI?.Log($"  [{cnx}] Warning: no DXF file found for key '{cnx}'. Skipped.");
                    skipped += pairIndices.Count;
                    continue;
                }

                // Pre-parse DXF once — reused for every pair in this group
                List<XYZ> groupRawPoints = null;

                try
                {
                    groupRawPoints = ParseDxfPoints(dxfPath);
                }
                catch (Exception ex)
                {
                    UI?.Log($"  [{cnx}] DXF read error: {ex.Message}. Skipped.");
                    skipped += pairIndices.Count;
                    continue;
                }

                if (groupRawPoints == null || groupRawPoints.Count < 2)
                {
                    UI?.Log($"  [{cnx}] Error: DXF has insufficient points.");
                    skipped += pairIndices.Count;
                    continue;
                }

                foreach (var pair in pairIndices)
                {
                    var ptA = pts[pair.iA];
                    var ptB = pts[pair.iB];

                    // ── Deterministic A/B: smallest X wins (then Y, then Z) ──
                    // Removes dependency on Revit's selection/sort order.
                    const double sortTol = 0.01;
                    bool swapAB = false;
                    if (Math.Abs(ptA.pos.X - ptB.pos.X) > sortTol)
                        swapAB = ptA.pos.X > ptB.pos.X;
                    else if (Math.Abs(ptA.pos.Y - ptB.pos.Y) > sortTol)
                        swapAB = ptA.pos.Y > ptB.pos.Y;
                    else
                        swapAB = ptA.pos.Z > ptB.pos.Z;

                    if (swapAB)
                    {
                        var tmp = ptA; ptA = ptB; ptB = tmp;
                    }

                    UI?.Log($"  [{cnx}] ptA id:{ptA.id.IntegerValue} eq:'{ptA.equipoId}' pos:({ptA.pos.X:F2},{ptA.pos.Y:F2},{ptA.pos.Z:F2})" +
                            (swapAB ? " (sorted)" : ""));
                    UI?.Log($"  [{cnx}] ptB id:{ptB.id.IntegerValue} eq:'{ptB.equipoId}' pos:({ptB.pos.X:F2},{ptB.pos.Y:F2},{ptB.pos.Z:F2})");

                    try
                    {
                        double ftToMm  = UnitUtils.ConvertFromInternalUnits(1.0, UnitTypeId.Millimeters);
                        double rvtDx   = (ptB.pos.X - ptA.pos.X) * ftToMm;
                        double rvtDy   = (ptB.pos.Y - ptA.pos.Y) * ftToMm;
                        double rvtVano     = Math.Sqrt(rvtDx * rvtDx + rvtDy * rvtDy);
                        double rvtDesnivel = Math.Abs((ptB.pos.Z - ptA.pos.Z) * ftToMm);

                        UI?.Log($"  [{cnx}] RVT → vano: {rvtVano:F1} mm  |  desnivel: {rvtDesnivel:F1} mm");

                        // ── Orientation via DXF vector dot product ──────────────
                        XYZ pointA = ptA.pos;
                        XYZ pointB = ptB.pos;

                        XYZ dxfStart = groupRawPoints[0];
                        XYZ dxfEnd = groupRawPoints[groupRawPoints.Count - 1];
                        XYZ dxfVec2D = new XYZ(dxfEnd.X - dxfStart.X, dxfEnd.Y - dxfStart.Y, 0);
                        XYZ rvtVec2D = new XYZ(pointB.X - pointA.X, pointB.Y - pointA.Y, 0);

                        double dot;
                        bool reverseDxf;

                        if (rvtDesnivel > rvtVano * 1.5)
                        {
                            // Desnivel-dominant: DXF Y = vertical, compare with Revit Z
                            double rvtDz = (ptB.pos.Z - ptA.pos.Z);  // in feet, signed
                            double dxfDy = dxfEnd.Y - dxfStart.Y;    // in feet, signed
                            dot = dxfDy * rvtDz;
                            reverseDxf = dot < 0;
                            UI?.Log($"  [{cnx}] Desnivel-dominant → using DXF.Y({dxfDy:F2}) × RVT.Z({rvtDz:F2})");
                        }
                        else
                        {
                            // Vano-dominant: standard 2D XY dot product
                            dot = dxfVec2D.DotProduct(rvtVec2D);
                            reverseDxf = dot < 0;
                        }

                        UI?.Log($"  [{cnx}] DXF vec:({dxfVec2D.X:F2},{dxfVec2D.Y:F2}) RVT vec:({rvtVec2D.X:F2},{rvtVec2D.Y:F2}) dot:{dot:F2} → reverse={reverseDxf}");
                        double dxfVanoMm     = Math.Abs(dxfEnd.X - dxfStart.X) * ftToMm;
                        double dxfDesnivelMm = Math.Abs(dxfEnd.Y - dxfStart.Y) * ftToMm;
                        UI?.Log($"  [{cnx}] DXF → vano: {dxfVanoMm:F1} mm  |  desnivel: {dxfDesnivelMm:F1} mm");

                        List<XYZ> dxfPoints = reverseDxf
                            ? Enumerable.Reverse(groupRawPoints).ToList()
                            : groupRawPoints;

                       
                        List<XYZ> trajectory = CalculateTrajectory(dxfPoints, pointA, pointB, false);
                        trajectory = ResampleTrajectory(trajectory, 0.5);

                        double minClearance = 0.5;
                        while (trajectory.Count > 0 && trajectory.First().DistanceTo(pointA) <= minClearance)
                            trajectory.RemoveAt(0);
                        while (trajectory.Count > 0 && trajectory.Last().DistanceTo(pointB) <= minClearance)
                            trajectory.RemoveAt(trajectory.Count - 1);

                        trajectory.Insert(0, pointA);
                        trajectory.Add(pointB);

                        XYZ startTangent = (trajectory[1] - trajectory[0]).Normalize();
                        XYZ endTangent   = (trajectory[trajectory.Count - 1] - trajectory[trajectory.Count - 2]).Normalize();

                        using (Transaction trans = new Transaction(doc, "HMV - Create FlexPipe " + cnx))
                        {
                            trans.Start();

                            FailureHandlingOptions fho = trans.GetFailureHandlingOptions();
                            fho.SetFailuresPreprocessor(new WarningSuppressor());
                            trans.SetFailureHandlingOptions(fho);

                            FlexPipe flexPipe = FlexPipe.Create(
                                doc, systemTypeId, flexType.Id, defaultLevel.Id,
                                startTangent, endTangent, trajectory);

                            if (flexPipe == null)
                            {
                                trans.RollBack();
                                UI?.Log($"  [{cnx}] Error: FlexPipe.Create returned null.");
                                skipped++;
                                continue;
                            }

                            // Inicial/Final based on DXF vector decision
                            var ptInicial = reverseDxf ? ptB : ptA;
                            var ptFinal = reverseDxf ? ptA : ptB;

                            Parameter diamParam = flexPipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                            if (diamParam != null && !diamParam.IsReadOnly)
                                diamParam.Set(UnitUtils.ConvertToInternalUnits(0.75, UnitTypeId.Inches));

                            TrySetStringParam(flexPipe, Config.FlexCnxNumberParam,    cnx);
                            TrySetStringParam(flexPipe, Config.FlexEquipoInicialParam, StripEquipmentId(ptInicial.equipoId));
                            TrySetStringParam(flexPipe, Config.FlexEquipoFinalParam,   StripEquipmentId(ptFinal.equipoId));

                            SetConnectedFlag(ptInicial.fi, Config.ConnectedParam);
                            SetConnectedFlag(ptFinal.fi,   Config.ConnectedParam);

                            trans.Commit();

                            UI?.Log($"  [{cnx}] ✔ Created — Inicial: {ptInicial.equipoId} → Final: {ptFinal.equipoId}" +
                                    (reverseDxf ? " (reversed)" : ""));
                        }
                        created++;
                    }
                    catch (Exception ex)
                    {
                        UI?.Log($"  [{cnx}] Exception: {ex.Message}");
                        skipped++;
                    }
                }
            }

            UI?.Log($"── Done: {created} created, {skipped} skipped ─────────────");
            UI?.SetStatus($"✔ {created} FlexPipe(s) created, {skipped} skipped.");
        }

        // ═══════════════════════════════════════════════════════════════════
        //  AUTO-ORIENTATION — deviation profile fingerprint
        // ═══════════════════════════════════════════════════════════════════

        /// <summary>
        /// Samples the UNSIGNED perpendicular deviation of a polyline from its chord
        /// at uniform arc-length fractions. Unsigned so it's comparable across
        /// different coordinate systems (DXF space vs Revit space).
        /// </summary>
        private static double[] ComputeDeviationProfileUnsigned(
            IList<XYZ> polyline, XYZ chordStart, XYZ chordEnd, int sampleCount)
        {
            XYZ    chordVec = chordEnd - chordStart;
            double chordLen = chordVec.GetLength();
            var    profile  = new double[sampleCount];

            if (chordLen < 1e-9 || polyline.Count < 2) return profile;

            XYZ uAB = chordVec.Normalize();

            // Build cumulative arc-length table
            var arcLen = new double[polyline.Count];
            arcLen[0] = 0;
            for (int i = 1; i < polyline.Count; i++)
                arcLen[i] = arcLen[i - 1] + polyline[i].DistanceTo(polyline[i - 1]);

            double totalArc = arcLen[polyline.Count - 1];
            if (totalArc < 1e-9) return profile;

            for (int s = 0; s < sampleCount; s++)
            {
                double targetArc = totalArc * (s + 1) / (sampleCount + 1);

                // Find segment
                int seg = 1;
                while (seg < polyline.Count - 1 && arcLen[seg] < targetArc) seg++;

                double segLen = arcLen[seg] - arcLen[seg - 1];
                double segFrac = segLen > 1e-9
                    ? (targetArc - arcLen[seg - 1]) / segLen
                    : 0;

                XYZ pt = polyline[seg - 1] + (polyline[seg] - polyline[seg - 1]) * segFrac;

                // Unsigned perpendicular distance to chord
                XYZ local     = pt - chordStart;
                double along  = local.DotProduct(uAB);
                XYZ projected = chordStart + uAB * along;
                profile[s]    = pt.DistanceTo(projected);  // unsigned!
            }

            return profile;
        }

        /// <summary>
        /// Returns true if reversing the DXF point order would better match the Revit curve.
        /// Sets <paramref name="isSymmetric"/> = true when forward and reverse scores are
        /// within 1 % of each other (ambiguous — no auto-correction applied).
        /// </summary>
        private static bool CheckNeedsReverse(
            double[] dxfProfile, double[] rvtProfile, out bool isSymmetric)
        {
            isSymmetric = false;
            int n = Math.Min(dxfProfile.Length, rvtProfile.Length);
            if (n == 0) return false;

            double forwardScore = 0;
            double reverseScore = 0;

            for (int i = 0; i < n; i++)
            {
                forwardScore += dxfProfile[i] * rvtProfile[i];
                reverseScore += dxfProfile[n - 1 - i] * rvtProfile[i];
            }

            double magnitude = Math.Max(Math.Abs(forwardScore), Math.Abs(reverseScore));
            if (magnitude > 1e-9 &&
                Math.Abs(forwardScore - reverseScore) / magnitude < 0.01)
            {
                isSymmetric = true;
                return false;
            }

            return reverseScore > forwardScore;
        }

        // ═══════════════════════════════════════════════════════════════════
        //  DXF HELPERS
        // ═══════════════════════════════════════════════════════════════════

        // ── Full DXF point parsing — mirrors ElectricalConnectionCommand.ParseDxfPoints ──

        private static List<XYZ> ParseDxfPoints(string filePath)
        {
            DxfFile dxf = DxfFile.Load(filePath);
            var points = new List<XYZ>();

            foreach (DxfEntity entity in dxf.Entities)
            {
                switch (entity)
                {
                    case DxfLwPolyline lw:
                        foreach (DxfLwPolylineVertex v in lw.Vertices)
                            points.Add(MmToFeet(v.X, v.Y, 0));
                        break;

                    case DxfPolyline poly:
                        foreach (DxfVertex v in poly.Vertices)
                            points.Add(MmToFeet(v.Location.X, v.Location.Y, v.Location.Z));
                        break;

                    case DxfLine line:
                        if (points.Count == 0)
                            points.Add(MmToFeet(line.P1.X, line.P1.Y, line.P1.Z));
                        points.Add(MmToFeet(line.P2.X, line.P2.Y, line.P2.Z));
                        break;

                    case DxfSpline spline:
                        foreach (DxfControlPoint cp in spline.ControlPoints)
                            points.Add(MmToFeet(cp.Point.X, cp.Point.Y, cp.Point.Z));
                        break;
                }
            }

            if (points.Count < 2) return null;

            var cleaned = new List<XYZ> { points[0] };
            for (int i = 1; i < points.Count; i++)
                if (points[i].DistanceTo(cleaned[cleaned.Count - 1]) > 0.001)
                    cleaned.Add(points[i]);

            return cleaned;
        }

        private static XYZ MmToFeet(double x, double y, double z)
        {
            return new XYZ(
                UnitUtils.ConvertToInternalUnits(x, UnitTypeId.Millimeters),
                UnitUtils.ConvertToInternalUnits(y, UnitTypeId.Millimeters),
                UnitUtils.ConvertToInternalUnits(z, UnitTypeId.Millimeters));
        }

        // ── Trajectory — mirrors ElectricalConnectionCommand.CalculateTrajectory ──

        private static List<XYZ> CalculateTrajectory(List<XYZ> dxfPoints, XYZ originA, XYZ originB, bool mirror = false)
        {
            XYZ dxfStart = dxfPoints[0];
            XYZ dxfEnd   = dxfPoints[dxfPoints.Count - 1];
            XYZ dxfVec   = dxfEnd - dxfStart;
            double dxfLen = dxfVec.GetLength();

            if (dxfLen < 1e-9) throw new InvalidOperationException("DXF start and end are identical.");

            XYZ uDxf = dxfVec.Normalize();
            XYZ vDxf = XYZ.BasisZ.CrossProduct(uDxf).Normalize();

            XYZ revVec = originB - originA;
            double revLen = revVec.GetLength();
            if (revLen < 1e-9) throw new InvalidOperationException("Points A and B are at the same position.");

            XYZ uRev = revVec.Normalize();
            XYZ horizontalNormal = uRev.CrossProduct(XYZ.BasisZ);
            XYZ vRev = horizontalNormal.GetLength() < 1e-9
                ? XYZ.BasisX
                : horizontalNormal.Normalize().CrossProduct(uRev).Normalize();

            double scale       = revLen / dxfLen;
            double deviateSign = mirror ? -1.0 : 1.0;

            var transformed = new List<XYZ>();
            foreach (XYZ pt in dxfPoints)
            {
                XYZ    local   = pt - dxfStart;
                double along   = local.DotProduct(uDxf);
                double deviate = local.DotProduct(vDxf);
                transformed.Add(originA + (uRev * along * scale) + (vRev * deviate * deviateSign * scale));
            }

            transformed[transformed.Count - 1] = originB;
            return transformed;
        }

        private static List<XYZ> ResampleTrajectory(List<XYZ> points, double interval = 0.3)
        {
            double totalLen = 0;
            for (int i = 1; i < points.Count; i++)
                totalLen += points[i].DistanceTo(points[i - 1]);

            if (totalLen < interval * 2) return points;

            var result = new List<XYZ> { points[0] };
            double sinceLastSample = 0;

            for (int i = 1; i < points.Count; i++)
            {
                double segLen = points[i].DistanceTo(points[i - 1]);
                XYZ dir = (points[i] - points[i - 1]).Normalize();
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

        // ═══════════════════════════════════════════════════════════════════
        //  TYPE RESOLVERS
        // ═══════════════════════════════════════════════════════════════════

        private static FlexPipeType FindFlexPipeType(Document doc, string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return null;

            string[] parts = key.Split(new[] { " : " }, StringSplitOptions.None);
            return parts.Length == 2
                ? new FilteredElementCollector(doc).OfClass(typeof(FlexPipeType))
                    .Cast<FlexPipeType>()
                    .FirstOrDefault(ft => ft.FamilyName == parts[0] && ft.Name == parts[1])
                : new FilteredElementCollector(doc).OfClass(typeof(FlexPipeType))
                    .Cast<FlexPipeType>()
                    .FirstOrDefault(ft => ft.Name == key);
        }

        private static ElementId FindElectricalSystemTypeId(Document doc, string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return null;

            var st = new FilteredElementCollector(doc)
                .OfClass(typeof(PipingSystemType))
                .Cast<PipingSystemType>()
                .FirstOrDefault(s => s.Name.Equals(key, StringComparison.OrdinalIgnoreCase));

            return st?.Id;
        }

        // ═══════════════════════════════════════════════════════════════════
        //  SMALL UTILITIES
        // ═══════════════════════════════════════════════════════════════════

        /// <summary>
        /// Returns greedy minimum-distance index pairs from a list of candidate points.
        /// Pairs where both points share the same equipment ID are excluded.
        /// </summary>
        private static List<(int iA, int iB)> FindOptimalPairIndices(
            List<(ElementId id, FamilyInstance fi, string cnx, string equipoId, XYZ pos)> pts)
        {
            var result = new List<(int, int)>();
            var used   = new HashSet<int>();

            var candidates = new List<(int iA, int iB, double dist)>();
            for (int i = 0; i < pts.Count; i++)
            {
                for (int j = i + 1; j < pts.Count; j++)
                {
                    string idA = ExtractEquipmentId(pts[i].equipoId);
                    string idB = ExtractEquipmentId(pts[j].equipoId);
                    if (!string.IsNullOrWhiteSpace(idA) &&
                        idA.Equals(idB, StringComparison.OrdinalIgnoreCase))
                        continue;
                    candidates.Add((i, j, pts[i].pos.DistanceTo(pts[j].pos)));
                }
            }

            candidates.Sort((a, b) => a.dist.CompareTo(b.dist));

            foreach (var c in candidates)
            {
                if (used.Contains(c.iA) || used.Contains(c.iB)) continue;
                result.Add((c.iA, c.iB));
                used.Add(c.iA);
                used.Add(c.iB);
            }

            return result;
        }

        /// <summary>
        /// Extracts the equipment ElementId portion from a value formatted as "EQUIPO_ID".
        /// "TC_2412545" → "2412545".  If no '_' is found, returns the full string.
        /// </summary>
        private static string ExtractEquipmentId(string equipoValue)
        {
            if (string.IsNullOrWhiteSpace(equipoValue)) return string.Empty;
            int idx = equipoValue.LastIndexOf('_');
            return idx >= 0 ? equipoValue.Substring(idx + 1) : equipoValue;
        }

        /// <summary>
        /// Returns only the equipment name portion, stripping the "_ElementId" suffix.
        /// "TC_2412545" → "TC".  If no '_' is found, returns the full string.
        /// </summary>
        private static string StripEquipmentId(string equipoValue)
        {
            if (string.IsNullOrWhiteSpace(equipoValue)) return string.Empty;
            int idx = equipoValue.LastIndexOf('_');
            return idx >= 0 ? equipoValue.Substring(0, idx) : equipoValue;
        }

        private static Dictionary<string, string> BuildDxfMap(string folder)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string path in Directory.GetFiles(folder, "*.dxf", SearchOption.AllDirectories))
            {
                string name = Path.GetFileNameWithoutExtension(path);
                int idx = name.IndexOf('_');
                string key = idx > 0 ? name.Substring(0, idx) : name;
                if (!map.ContainsKey(key)) map[key] = path;
            }
            return map;
        }

        private static XYZ GetAdaptivePosition(Document doc, FamilyInstance fi)
        {
            try
            {
                IList<ElementId> ptIds =
                    AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(fi);
                if (ptIds != null && ptIds.Count > 0)
                {
                    ReferencePoint rp = doc.GetElement(ptIds[0]) as ReferencePoint;
                    if (rp != null) return rp.Position;
                }
            }
            catch { }
            return (fi.Location as LocationPoint)?.Point ?? XYZ.Zero;
        }

        private static string GetStringParam(Element elem, string paramName)
        {
            if (elem == null || string.IsNullOrWhiteSpace(paramName)) return string.Empty;
            Parameter p = elem.LookupParameter(paramName);
            if (p == null) return string.Empty;
            return p.StorageType == StorageType.String ? p.AsString() ?? string.Empty : string.Empty;
        }

        private static void TrySetStringParam(Element elem, string paramName, string value)
        {
            if (elem == null || string.IsNullOrWhiteSpace(paramName)) return;
            Parameter p = elem.LookupParameter(paramName);
            if (p == null || p.IsReadOnly || p.StorageType != StorageType.String) return;
            p.Set(value ?? string.Empty);
        }

        private static void SetConnectedFlag(FamilyInstance fi, string paramName)
        {
            if (fi == null || string.IsNullOrWhiteSpace(paramName)) return;
            Parameter p = fi.LookupParameter(paramName);
            if (p == null || p.IsReadOnly || p.StorageType != StorageType.Integer) return;
            p.Set(1);
        }
    }

    // ═══════════════════════════════════════════════════════════════════
    //  MIRROR SUPPORT
    // ═══════════════════════════════════════════════════════════════════

    /// <summary>
    /// Picks one or more FlexPipes and mirrors each about its own A→B axis,
    /// flipping the bulge to the opposite side. Deletes the originals.
    /// </summary>
    public class MirrorFlexPipeRefresherHandler : IExternalEventHandler
    {
        public ElectricalRefresherWindow UI { get; set; }

        public string GetName() => "MirrorFlexPipeRefresherHandler";

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document   doc   = uidoc.Document;

            try
            {
                // ── 1. Pick one or more FlexPipes ────────────────────────────
                IList<Reference> refs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    new FlexPipeSelectionFilter(),
                    "Select FlexPipe(s) to mirror — click Finish when done");

                if (refs == null || refs.Count == 0)
                {
                    UI?.SetStatus("No FlexPipes selected.");
                    UI?.RestoreWindow();
                    return;
                }

                int mirrored = 0;
                int failed   = 0;

                using (Transaction t = new Transaction(doc, "HMV - Mirror FlexPipe(s)"))
                {
                    t.Start();

                    FailureHandlingOptions fho = t.GetFailureHandlingOptions();
                    fho.SetFailuresPreprocessor(new WarningSuppressor());
                    t.SetFailureHandlingOptions(fho);

                    foreach (Reference elemRef in refs)
                    {
                        FlexPipe flexPipe = doc.GetElement(elemRef) as FlexPipe;
                        if (flexPipe == null) { failed++; continue; }

                        // ── 2. Get trajectory points ──────────────────────────
                        IList<XYZ> pts = flexPipe.Points;
                        if (pts == null || pts.Count < 3) { failed++; continue; }

                        XYZ ptA = pts[0];
                        XYZ ptB = pts[pts.Count - 1];
                        XYZ uAB = (ptB - ptA).Normalize();

                        // ── 3. Find the maximum-deviation (bulge) point ────────
                        XYZ bulgePt  = null;
                        double maxDev = 0.0;

                        for (int i = 1; i < pts.Count - 1; i++)
                        {
                            XYZ local     = pts[i] - ptA;
                            double along  = local.DotProduct(uAB);
                            XYZ projected = ptA + uAB * along;
                            double dev    = pts[i].DistanceTo(projected);
                            if (dev > maxDev) { maxDev = dev; bulgePt = pts[i]; }
                        }

                        if (bulgePt == null || maxDev < 1e-6) { failed++; continue; }

                        // ── 4. Build mirror plane ──────────────────────────────
                        XYZ local2     = bulgePt - ptA;
                        double along2  = local2.DotProduct(uAB);
                        XYZ projected2 = ptA + uAB * along2;
                        XYZ bulgeDir   = (bulgePt - projected2).Normalize();

                        Plane mirrorPlane = Plane.CreateByNormalAndOrigin(bulgeDir, ptA);

                        // ── 5. Mirror + delete original ────────────────────────
                        ElementTransformUtils.MirrorElement(doc, flexPipe.Id, mirrorPlane);
                        doc.Delete(flexPipe.Id);

                        mirrored++;
                    }

                    t.Commit();
                }

                UI?.SetStatus($"✔ {mirrored} FlexPipe(s) mirrored" +
                              (failed > 0 ? $", {failed} skipped (too few points)." : "."));
                UI?.RestoreWindow();
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                UI?.SetStatus("Mirror canceled by user.");
                UI?.RestoreWindow();
            }
            catch (Exception ex)
            {
                UI?.SetStatus("Mirror error: " + ex.Message);
                UI?.RestoreWindow();
            }
        }
    }
}
