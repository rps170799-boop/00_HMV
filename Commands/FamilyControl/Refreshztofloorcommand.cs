using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class RefreshZToFloorCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // ── Step 1: Gather pre-selected or prompt-selected elements ──

            IList<Reference> elemRefs;
            ICollection<ElementId> preSelected = uidoc.Selection.GetElementIds();

            if (preSelected.Count > 0)
            {
                // Convert pre-selection to References
                elemRefs = preSelected
                    .Select(id => new Reference(doc.GetElement(id)))
                    .ToList();
            }
            else
            {
                try
                {
                    elemRefs = uidoc.Selection.PickObjects(
                        ObjectType.Element,
                        "Select foundation elements to refresh Z");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }
            }

            if (elemRefs == null || elemRefs.Count == 0)
            {
                TaskDialog.Show("HMV Tools", "No elements selected.");
                return Result.Cancelled;
            }

            // Validate: keep only elements with a LocationPoint
            var validElements = new List<Element>();
            foreach (var r in elemRefs)
            {
                Element el = doc.GetElement(r);
                if (el != null && el.Location is LocationPoint)
                    validElements.Add(el);
            }

            if (validElements.Count == 0)
            {
                TaskDialog.Show("HMV Tools",
                    "None of the selected elements have a point-based location.");
                return Result.Failed;
            }

            // ── Step 2: Collect available Revit links ────────────────

            var linkInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(li => li.GetLinkDocument() != null)
                .ToList();

            if (linkInstances.Count == 0)
            {
                TaskDialog.Show("HMV Tools",
                    "No loaded Revit links found in the project.");
                return Result.Failed;
            }

            // Build entries for the window
            var linkEntries = linkInstances.Select(li => new LinkPickEntry
            {
                Name = li.GetLinkDocument().Title,
                LinkId = li.Id.IntegerValue
            }).ToList();

            // ── Step 3: Show window ──────────────────────────────────

            var win = new RefreshZToFloorWindow(
                validElements.Count, linkEntries);

            bool? result = win.ShowDialog();
            if (result != true || win.SelectedLinkId < 0)
                return Result.Cancelled;

            // Resolve the chosen link
            RevitLinkInstance chosenLink = doc.GetElement(
                new ElementId(win.SelectedLinkId)) as RevitLinkInstance;

            if (chosenLink == null)
            {
                TaskDialog.Show("HMV Tools", "Could not resolve the selected link.");
                return Result.Failed;
            }

            Document linkDoc = chosenLink.GetLinkDocument();
            Transform linkTx = chosenLink.GetTotalTransform();

            // ── Step 4: Get floor solids in host coordinates ─────────

            var floorSolids = CollectFloorSolids(linkDoc, linkTx);

            if (floorSolids.Count == 0)
            {
                TaskDialog.Show("HMV Tools",
                    "No floor geometry found in the selected link.");
                return Result.Failed;
            }

            // ── Step 5: Ray-cast and move elements ───────────────────

            int movedCount = 0;
            int skippedCount = 0;
            int unchangedCount = 0;
            var reportLines = new List<string>();

            using (Transaction tx = new Transaction(doc, "Refresh Z to Linked Floor"))
            {
                tx.Start();

                foreach (Element elem in validElements)
                {
                    LocationPoint locPt = elem.Location as LocationPoint;
                    XYZ current = locPt.Point;

                    // Vertical ray through element X,Y
                    double? topZ = FindTopFloorZ(
                        current.X, current.Y, current.Z, floorSolids);

                    if (!topZ.HasValue)
                    {
                        skippedCount++;
                        reportLines.Add(
                            $"SKIP  │ {elem.Name}  (Id {elem.Id.IntegerValue})"
                            + $"  — No floor intersection at X,Y");
                        continue;
                    }

                    double deltaZ = topZ.Value - current.Z;

                    if (Math.Abs(deltaZ) < 1e-9)
                    {
                        unchangedCount++;
                        reportLines.Add(
                            $"OK    │ {elem.Name}  (Id {elem.Id.IntegerValue})"
                            + $"  — Already at correct Z");
                        continue;
                    }

                    ElementTransformUtils.MoveElement(
                        doc, elem.Id, new XYZ(0, 0, deltaZ));

                    movedCount++;
                    double deltaZmm = deltaZ * 304.8;
                    reportLines.Add(
                        $"MOVED │ {elem.Name}  (Id {elem.Id.IntegerValue})"
                        + $"  — ΔZ = {deltaZmm:F1} mm");
                }

                tx.Commit();
            }

            // ── Step 6: Show report ──────────────────────────────────

            string summary =
                $"Elements moved: {movedCount}\n"
                + $"Already correct: {unchangedCount}\n"
                + $"Skipped (no intersection): {skippedCount}";

            var reportWin = new RefreshZReportWindow(summary, reportLines);
            reportWin.ShowDialog();

            return Result.Succeeded;
        }

        // ═══════════════════════════════════════════════════════════
        //  Helpers
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// Collects all floor solids from the linked document,
        /// transformed into the host model coordinate system.
        /// </summary>
        private List<Solid> CollectFloorSolids(
            Document linkDoc, Transform linkTx)
        {
            var solids = new List<Solid>();
            Options opts = new Options
            {
                ComputeReferences = false,
                DetailLevel = ViewDetailLevel.Fine
            };

            var floors = new FilteredElementCollector(linkDoc)
                .OfCategory(BuiltInCategory.OST_Floors)
                .WhereElementIsNotElementType()
                .ToList();

            foreach (var floor in floors)
            {
                GeometryElement geomElem = floor.get_Geometry(opts);
                if (geomElem == null) continue;

                foreach (GeometryObject geomObj in geomElem)
                {
                    Solid solid = geomObj as Solid;
                    if (solid == null || solid.Volume < 1e-9) continue;

                    Solid transformed =
                        SolidUtils.CreateTransformed(solid, linkTx);
                    solids.Add(transformed);
                }
            }

            return solids;
        }

        /// <summary>
        /// Shoots a vertical ray at (x, y) and finds the highest
        /// floor-top-face Z that is closest to the element's current Z.
        /// Prioritises the top face (upward-pointing normal) and
        /// picks the intersection nearest to currentZ.
        /// </summary>
        private double? FindTopFloorZ(
            double x, double y, double currentZ,
            List<Solid> floorSolids)
        {
            double searchSpan = 500;  // feet – generous vertical range
            XYZ top = new XYZ(x, y, currentZ + searchSpan);
            XYZ bot = new XYZ(x, y, currentZ - searchSpan);
            Line ray = Line.CreateBound(top, bot);

            // Collect all intersection Z values on upward faces
            var hitZValues = new List<double>();

            foreach (Solid solid in floorSolids)
            {
                foreach (Face face in solid.Faces)
                {
                    // Keep only top faces: normal Z component > 0
                    // (using the face normal at UV origin as a proxy)
                    BoundingBoxUV bbUV = face.GetBoundingBox();
                    UV midUV = new UV(
                        (bbUV.Min.U + bbUV.Max.U) / 2.0,
                        (bbUV.Min.V + bbUV.Max.V) / 2.0);
                    XYZ normal = face.ComputeNormal(midUV);

                    if (normal.Z < 0.3) continue; // skip bottom/side faces

                    IntersectionResultArray results;
                    SetComparisonResult cmp =
                        face.Intersect(ray, out results);

                    if (cmp != SetComparisonResult.Overlap
                        || results == null) continue;

                    foreach (IntersectionResult ir in results)
                        hitZValues.Add(ir.XYZPoint.Z);
                }
            }

            if (hitZValues.Count == 0) return null;

            // Pick the hit closest to the element's current Z
            return hitZValues
                .OrderBy(z => Math.Abs(z - currentZ))
                .First();
        }
    }

    // ── Selection filter: RevitLinkInstance only ────────────────

    public class LinkOnlySelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is RevitLinkInstance;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }

    // ── Plain data class for the window ────────────────────────

    public class LinkPickEntry
    {
        public string Name { get; set; }
        public int LinkId { get; set; }
    }
}