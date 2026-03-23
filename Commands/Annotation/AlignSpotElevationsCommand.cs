using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class AlignSpotElevationsCommand : IExternalCommand
    {
        private class SpotElevationFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            { return elem is SpotDimension; }
            public bool AllowReference(Reference reference, XYZ position)
            { return false; }
        }

        private class CurveElementFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            { return elem is CurveElement; }
            public bool AllowReference(Reference reference, XYZ position)
            { return false; }
        }

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // ── Step 1: Collect spot elevations ─────────────────
            var spots = new List<SpotDimension>();

            ICollection<ElementId> preSelected =
                uidoc.Selection.GetElementIds();

            if (preSelected != null && preSelected.Count > 0)
            {
                foreach (ElementId id in preSelected)
                {
                    var spot = doc.GetElement(id) as SpotDimension;
                    if (spot != null) spots.Add(spot);
                }
            }

            if (spots.Count == 0)
            {
                IList<Reference> pickedRefs;
                try
                {
                    pickedRefs = uidoc.Selection.PickObjects(
                        ObjectType.Element,
                        new SpotElevationFilter(),
                        "Select Spot Elevations to align. Press Finish when done.");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                { return Result.Cancelled; }

                if (pickedRefs == null || pickedRefs.Count == 0)
                    return Result.Cancelled;

                foreach (Reference r in pickedRefs)
                {
                    var spot = doc.GetElement(r) as SpotDimension;
                    if (spot != null) spots.Add(spot);
                }
            }

            if (spots.Count == 0)
            {
                TaskDialog.Show("HMV Tools",
                    "No spot elevations found in selection.");
                return Result.Cancelled;
            }

            // ── Step 2: Show settings window ────────────────────
            var win = new SpotAlignmentWindow(spots.Count);
            if (win.ShowDialog() != true || win.Settings == null)
                return Result.Cancelled;

            bool moveWithLeader = win.Settings.MoveWithLeader;

            // ── Step 3: Pick reference line ──────────────────────
            Reference lineRef;
            try
            {
                lineRef = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new CurveElementFilter(),
                    "Pick a reference line (model or detail line).");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            { return Result.Cancelled; }

            if (lineRef == null) return Result.Cancelled;

            var curveElem = doc.GetElement(lineRef) as CurveElement;
            if (curveElem == null || curveElem.GeometryCurve == null)
            {
                TaskDialog.Show("HMV Tools",
                    "Could not read geometry from the selected line.");
                return Result.Cancelled;
            }

            // ── Auto-detect axis from line orientation ──────────
            XYZ p0 = curveElem.GeometryCurve.GetEndPoint(0);
            XYZ p1 = curveElem.GeometryCurve.GetEndPoint(1);
            double dx = Math.Abs(p1.X - p0.X);
            double dy = Math.Abs(p1.Y - p0.Y);
            // If line is more horizontal (dx > dy), it defines a Y coord → align Y
            // If line is more vertical (dy > dx), it defines an X coord → align X
            bool alignX = dy >= dx;
            double targetCoord = alignX ? p0.X : p0.Y;

            // ── Step 4: Align spots ─────────────────────────────
            int aligned = 0;
            var errors = new List<string>();
            var debug = new List<string>();

            debug.Add($"Axis: {(alignX ? "X" : "Y")}, Target: {targetCoord:F4} ft");
            debug.Add($"Mode: {(moveWithLeader ? "Leader+Text" : "Text only")}");

            using (Transaction tx = new Transaction(doc,
                "HMV – Align Spot Elevations"))
            {
                tx.Start();

                foreach (SpotDimension spot in spots)
                {
                    try
                    {
                        if (moveWithLeader)
                        {
                            // ════════════════════════════════════
                            // LEADER MODE: move LeaderShoulderPosition
                            // Text follows the shoulder automatically
                            // ════════════════════════════════════

                            // Try LeaderShoulderPosition first
                            bool moved = false;

                            try
                            {
                                XYZ oldShoulder = spot.LeaderShoulderPosition;
                                if (oldShoulder != null)
                                {
                                    XYZ newShoulder = alignX
                                        ? new XYZ(targetCoord, oldShoulder.Y, oldShoulder.Z)
                                        : new XYZ(oldShoulder.X, targetCoord, oldShoulder.Z);

                                    spot.LeaderShoulderPosition = newShoulder;

                                    // Verify
                                    XYZ check = spot.LeaderShoulderPosition;
                                    moved = check != null
                                        && (alignX
                                            ? Math.Abs(check.X - targetCoord) < 0.01
                                            : Math.Abs(check.Y - targetCoord) < 0.01);
                                }
                            }
                            catch { }

                            if (moved)
                            {
                                aligned++;
                                continue;
                            }

                            // Fallback: try LeaderEndPosition
                            try
                            {
                                XYZ oldEnd = spot.LeaderEndPosition;
                                if (oldEnd != null)
                                {
                                    XYZ newEnd = alignX
                                        ? new XYZ(targetCoord, oldEnd.Y, oldEnd.Z)
                                        : new XYZ(oldEnd.X, targetCoord, oldEnd.Z);

                                    spot.LeaderEndPosition = newEnd;

                                    XYZ check = spot.LeaderEndPosition;
                                    moved = check != null
                                        && (alignX
                                            ? Math.Abs(check.X - targetCoord) < 0.01
                                            : Math.Abs(check.Y - targetCoord) < 0.01);
                                }
                            }
                            catch { }

                            if (moved)
                            {
                                aligned++;
                                continue;
                            }

                            // Last resort: TextPosition (moves text,
                            // leader stretches to follow)
                            try
                            {
                                XYZ oldTP = spot.TextPosition;
                                if (oldTP != null)
                                {
                                    XYZ newTP = alignX
                                        ? new XYZ(targetCoord, oldTP.Y, oldTP.Z)
                                        : new XYZ(oldTP.X, targetCoord, oldTP.Z);

                                    spot.TextPosition = newTP;

                                    XYZ check = spot.TextPosition;
                                    moved = check != null
                                        && (alignX
                                            ? Math.Abs(check.X - targetCoord) < 0.01
                                            : Math.Abs(check.Y - targetCoord) < 0.01);
                                }
                            }
                            catch { }

                            if (moved)
                                aligned++;
                            else
                                errors.Add(
                                    $"Id {spot.Id.IntegerValue}: "
                                  + "all leader approaches failed");
                        }
                        else
                        {
                            // ════════════════════════════════════
                            // TEXT-ONLY MODE (unchanged — working)
                            // ════════════════════════════════════

                            XYZ oldTP = null;
                            try { oldTP = spot.TextPosition; } catch { }

                            if (oldTP != null)
                            {
                                XYZ newTP = alignX
                                    ? new XYZ(targetCoord, oldTP.Y, oldTP.Z)
                                    : new XYZ(oldTP.X, targetCoord, oldTP.Z);

                                spot.TextPosition = newTP;

                                XYZ checkTP = spot.TextPosition;
                                bool movedTP = checkTP != null
                                    && (alignX
                                        ? Math.Abs(checkTP.X - targetCoord) < 0.01
                                        : Math.Abs(checkTP.Y - targetCoord) < 0.01);

                                if (movedTP) { aligned++; continue; }
                            }

                            // Segment fallback
                            try
                            {
                                if (spot.Segments.Size > 0)
                                {
                                    var seg = spot.Segments.get_Item(0);
                                    XYZ oldSTP = seg.TextPosition;
                                    if (oldSTP != null)
                                    {
                                        XYZ newSTP = alignX
                                            ? new XYZ(targetCoord, oldSTP.Y, oldSTP.Z)
                                            : new XYZ(oldSTP.X, targetCoord, oldSTP.Z);
                                        seg.TextPosition = newSTP;

                                        XYZ checkSTP = seg.TextPosition;
                                        bool movedSTP = checkSTP != null
                                            && (alignX
                                                ? Math.Abs(checkSTP.X - targetCoord) < 0.01
                                                : Math.Abs(checkSTP.Y - targetCoord) < 0.01);

                                        if (movedSTP) { aligned++; continue; }
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Id {spot.Id.IntegerValue}: {ex.Message}");
                    }
                }

                tx.Commit();
            }

            // ── Report ──────────────────────────────────────────
            string report =
                "ALIGN SPOT ELEVATIONS\n"
              + "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
              + $"Axis: {(alignX ? "X" : "Y")}  (auto-detected)\n"
              + $"Mode: {(moveWithLeader ? "Leader + Text" : "Text only")}\n"
              + $"Target: {targetCoord:F4} ft\n"
              + $"Spots aligned: {aligned} / {spots.Count}\n"
              + $"Errors: {errors.Count}\n";

            if (errors.Count > 0)
            {
                report += "\n── ERRORS ──\n";
                foreach (string e in errors)
                    report += $"  • {e}\n";
            }

            var td = new TaskDialog("HMV Tools – Align Spots");
            td.MainContent = report;
            td.Show();

            return Result.Succeeded;
        }
    }
}