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
    public class AutoDimensionCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // 1. Show UI to capture Tolerance
            var window = new AutoDimensionWindow();
            bool? result = window.ShowDialog();
            if (result != true)
            {
                return Result.Cancelled;
            }

            double toleranceCm = window.ToleranceCm;
            double toleranceInternal = toleranceCm / 30.48; // Convert CM to Revit Internal Units (Feet)

            try
            {
                // 2. Select the line that will dictate the dimension direction
                Reference lineRef = uidoc.Selection.PickObject(ObjectType.Element, new LineSelectionFilter(), "Select a Detail Line or Model Line to act as the dimension path.");
                CurveElement curveElement = doc.GetElement(lineRef) as CurveElement;

                if (curveElement == null || !(curveElement.GeometryCurve is Line))
                {
                    TaskDialog.Show("HMV Tools", "Selected element is not a straight line.");
                    return Result.Failed;
                }

                Line selectedLine = curveElement.GeometryCurve as Line;

                // 3. Find intersecting Wall Faces
                List<Tuple<Reference, XYZ, double>> intersectionPoints = new List<Tuple<Reference, XYZ, double>>();

                // Filter for walls strictly in the Active View (2D Context)
                var walls = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .OfClass(typeof(Wall))
                    .WhereElementIsNotElementType()
                    .ToElements();

                // Compute References is critical to be able to Dimension them
                Options geomOptions = new Options
                {
                    ComputeReferences = true,
                    IncludeNonVisibleObjects = true,
                    View = doc.ActiveView
                };

                // Raycast infinite line version to ensure faces crossing the line segment are captured
                Line unboundLine = Line.CreateUnbound(selectedLine.GetEndPoint(0), selectedLine.Direction);

                foreach (Element wall in walls)
                {
                    GeometryElement geomElem = wall.get_Geometry(geomOptions);
                    if (geomElem == null) continue;

                    foreach (GeometryObject geomObj in geomElem)
                    {
                        if (geomObj is Solid solid && solid.Faces.Size > 0)
                        {
                            foreach (Face face in solid.Faces)
                            {
                                if (face is PlanarFace pf)
                                {
                                    // 1. Skip horizontal faces (Floor/Ceiling limits of the wall)
                                    if (Math.Abs(pf.FaceNormal.Z) > 0.01) continue;

                                    // --- NEW FIX 1: PERPENDICULAR CHECK ---
                                    // The dimension line MUST be perpendicular to the face. 
                                    // If the dot product of the face normal and the line direction is ~1, they are perpendicular.
                                    double dotProduct = Math.Abs(pf.FaceNormal.DotProduct(selectedLine.Direction));
                                    if (dotProduct < 0.99) continue; // Skips walls running parallel to the line or at weird angles

                                    IntersectionResultArray results;
                                    SetComparisonResult intersectResult = pf.Intersect(unboundLine, out results);

                                    if (intersectResult == SetComparisonResult.Overlap && results != null && results.Size > 0)
                                    {
                                        IntersectionResult ir = results.get_Item(0);
                                        XYZ pt = ir.XYZPoint;

                                        // Ensure the intersection happened within the bounds of the line we drew
                                        double param = selectedLine.Project(pt).Parameter;
                                        if (param >= 0 && param <= 1)
                                        {
                                            // --- NEW FIX 2: PREVENT DUPLICATES ---
                                            // If walls are joined, they might yield overlapping faces at the exact same coordinate.
                                            // Check if we already registered an intersection at this exact point before adding it.
                                            if (!intersectionPoints.Any(x => x.Item2.IsAlmostEqualTo(pt)))
                                            {
                                                double dist = pt.DistanceTo(selectedLine.GetEndPoint(0));
                                                intersectionPoints.Add(new Tuple<Reference, XYZ, double>(pf.Reference, pt, dist));
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                if (intersectionPoints.Count < 2)
                {
                    TaskDialog.Show("HMV Tools", "Could not find enough intersecting wall faces along the selected line.");
                    return Result.Cancelled;
                }

                // 4. Sort points geometrically and apply Tolerance filter
                intersectionPoints = intersectionPoints.OrderBy(x => x.Item3).ToList();

                ReferenceArray refArray = new ReferenceArray();
                double lastDist = double.MinValue; // Tracks the last approved reference position

                foreach (var item in intersectionPoints)
                {
                    // If this is the first point, OR the distance from the last approved point is greater than the tolerance -> Add it
                    if (refArray.Size == 0 || (item.Item3 - lastDist) >= toleranceInternal)
                    {
                        refArray.Append(item.Item1);
                        lastDist = item.Item3;
                    }
                }

                // Check again post-tolerance filtering
                if (refArray.Size < 2)
                {
                    TaskDialog.Show("HMV Tools", "Not enough references left after applying the 2.5cm tolerance skip.");
                    return Result.Cancelled;
                }

                // 5. Build Dimension String in Transaction
                using (Transaction tx = new Transaction(doc, "HMV Auto Dimension Walls"))
                {
                    tx.Start();

                    // Create the dimension
                    Dimension newDim = doc.Create.NewDimension(doc.ActiveView, selectedLine, refArray);

                    // CRITICAL: We must regenerate the document so Revit builds the segments 
                    // internally before we try to move their text positions.
                    doc.Regenerate();

                    // --- NEW LOGIC: FIX TEXT COLLISIONS ---

                    // Define thresholds (in Revit internal units - feet)
                    // You might need to tweak these based on your View Scale and standard Text Size
                    double minSegmentLengthForText = 1.0; // E.g., if segment is less than 1 foot (~30cm), move text
                    double textOffsetDistance = 0.5;      // How far away to pull the text

                    if (newDim.Segments.Size > 0)
                    {
                        // Calculate the vector that points "Up" or "Down" relative to the dimension line on the screen
                        XYZ lineDir = selectedLine.Direction;
                        XYZ viewNormal = doc.ActiveView.ViewDirection;
                        XYZ perpendicularDir = lineDir.CrossProduct(viewNormal).Normalize();

                        bool alternateSide = true; // Used to zigzag the text up and down

                        foreach (DimensionSegment segment in newDim.Segments)
                        {
                            double segmentLength = segment.Value ?? 0;

                            // If the segment is too small, the text will likely overlap
                            if (segmentLength > 0 && segmentLength < minSegmentLengthForText)
                            {
                                // Grab current text position
                                XYZ currentTextPos = segment.TextPosition;

                                // Calculate new position (move it outwards)
                                XYZ offsetVector = perpendicularDir * (alternateSide ? textOffsetDistance : -textOffsetDistance);
                                XYZ newPos = currentTextPos + offsetVector;

                                try
                                {
                                    segment.TextPosition = newPos;
                                }
                                catch
                                {
                                    // Fallback in case Revit restricts the movement 
                                    // (e.g., if the dimension type strictly prohibits text movement)
                                }

                                // Flip the side for the next small segment so they don't stack on top of each other
                                alternateSide = !alternateSide;
                            }
                        }
                    }
                    else
                    {
                        // Handling for single-segment dimensions (rare for this script, but good practice)
                        double segmentLength = newDim.Value ?? 0;
                        if (segmentLength > 0 && segmentLength < minSegmentLengthForText)
                        {
                            XYZ lineDir = selectedLine.Direction;
                            XYZ viewNormal = doc.ActiveView.ViewDirection;
                            XYZ perpendicularDir = lineDir.CrossProduct(viewNormal).Normalize();

                            try
                            {
                                newDim.TextPosition = newDim.TextPosition + (perpendicularDir * textOffsetDistance);
                            }
                            catch { }
                        }
                    }

                    tx.Commit();
                }

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User pressed Esc during pick object
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }


    // Custom Selection Filter to restrict the PickObject solely to Lines
    public class LineSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is CurveElement curveElement && curveElement.GeometryCurve is Line;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}