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
        // ── Selection filters ──────────────────────────────────

        /// <summary>
        /// Allows only SpotDimension elements (elevations, coordinates, etc.).
        /// </summary>
        private class SpotElevationFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                return elem is SpotDimension;
            }
            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }

        /// <summary>
        /// Allows only CurveElements (model lines, detail lines).
        /// </summary>
        private class CurveElementFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                return elem is CurveElement;
            }
            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }

        // ── Execute ────────────────────────────────────────────

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // ── Step 1: Collect spot elevations ─────────────────
            var spots = new List<SpotDimension>();

            // Check pre-selection first
            ICollection<ElementId> preSelected =
                uidoc.Selection.GetElementIds();

            if (preSelected != null && preSelected.Count > 0)
            {
                foreach (ElementId id in preSelected)
                {
                    Element elem = doc.GetElement(id);
                    var spot = elem as SpotDimension;
                    if (spot != null)
                        spots.Add(spot);
                }
            }

            // If no valid spots from pre-selection, prompt user to pick
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
                {
                    return Result.Cancelled;
                }

                if (pickedRefs == null || pickedRefs.Count == 0)
                    return Result.Cancelled;

                foreach (Reference r in pickedRefs)
                {
                    var spot = doc.GetElement(r) as SpotDimension;
                    if (spot != null)
                        spots.Add(spot);
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

            SpotAlignmentSettings settings = win.Settings;
            bool alignX = settings.AlignToX;

            // ── Step 3: Determine target coordinate ─────────────
            double targetCoord;

            if (settings.PickReferenceLine)
            {
                // User picks a reference line on screen
                Reference lineRef;
                try
                {
                    lineRef = uidoc.Selection.PickObject(
                        ObjectType.Element,
                        new CurveElementFilter(),
                        "Pick a reference line (model or detail line).");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                if (lineRef == null)
                    return Result.Cancelled;

                var curveElem = doc.GetElement(lineRef) as CurveElement;
                if (curveElem == null || curveElem.GeometryCurve == null)
                {
                    TaskDialog.Show("HMV Tools",
                        "Could not read geometry from the selected line.");
                    return Result.Cancelled;
                }

                // Extract coordinate from the line's first endpoint.
                // For X-axis alignment we read X; for Y-axis we read Y.
                XYZ endPoint = curveElem.GeometryCurve
                    .GetEndPoint(0);
                targetCoord = alignX ? endPoint.X : endPoint.Y;
            }
            else
            {
                // User typed a coordinate value (already in feet)
                targetCoord = settings.TypedValue;
            }

            // ── Step 4: Align spots in a single transaction ─────
            int aligned = 0;
            int shoulderFixed = 0;
            var errors = new List<string>();

            using (Transaction tx = new Transaction(doc,
                "HMV – Align Spot Elevations"))
            {
                tx.Start();

                foreach (SpotDimension spot in spots)
                {
                    try
                    {
                        // Current text position
                        XYZ oldTextPos = spot.TextPosition;
                        if (oldTextPos == null) continue;

                        // Compute delta on the alignment axis
                        double delta;
                        XYZ newTextPos;

                        if (alignX)
                        {
                            // Align X coordinates (vertical stack)
                            delta = targetCoord - oldTextPos.X;
                            newTextPos = new XYZ(
                                targetCoord,
                                oldTextPos.Y,
                                oldTextPos.Z);
                        }
                        else
                        {
                            // Align Y coordinates (horizontal stack)
                            delta = targetCoord - oldTextPos.Y;
                            newTextPos = new XYZ(
                                oldTextPos.X,
                                targetCoord,
                                oldTextPos.Z);
                        }

                        // Move text position
                        spot.TextPosition = newTextPos;
                        aligned++;

                        // Move leader shoulder by the same delta so
                        // the leader line doesn't break at weird angles.
                        // Some spots have no shoulder → catch gracefully.
                        try
                        {
                            XYZ oldShoulder =
                                spot.LeaderShoulderPosition;
                            if (oldShoulder != null)
                            {
                                XYZ newShoulder;
                                if (alignX)
                                {
                                    newShoulder = new XYZ(
                                        oldShoulder.X + delta,
                                        oldShoulder.Y,
                                        oldShoulder.Z);
                                }
                                else
                                {
                                    newShoulder = new XYZ(
                                        oldShoulder.X,
                                        oldShoulder.Y + delta,
                                        oldShoulder.Z);
                                }
                                spot.LeaderShoulderPosition = newShoulder;
                                shoulderFixed++;
                            }
                        }
                        catch
                        {
                            // Spot has no leader shoulder – this is normal
                            // for spots without visible leaders.
                        }
                    }
                    catch (Exception ex)
                    {
                        string name = "Id " + spot.Id.IntegerValue;
                        errors.Add($"{name}: {ex.Message}");
                    }
                }

                tx.Commit();
            }

            // ── Report ──────────────────────────────────────────
            string axis = alignX ? "X" : "Y";
            string report =
                "ALIGN SPOT ELEVATIONS\n"
              + "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
              + $"Axis: {axis}\n"
              + $"Target coordinate: {targetCoord:F4} ft\n"
              + $"Spots aligned: {aligned}\n"
              + $"Shoulders adjusted: {shoulderFixed}\n"
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
