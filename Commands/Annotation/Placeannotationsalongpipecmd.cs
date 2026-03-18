using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class PlaceAnnotationsAlongPipeCmd : IExternalCommand
    {
        // ── UNIT HELPERS ───────────────────────────────────────────

        private static double MmToFeet(double mm) => mm / 304.8;

        // ── ENTRY POINT ────────────────────────────────────────────

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View activeView = doc.ActiveView;

            // ─ 1. Validate active view ─────────────────────────────
            if (activeView is View3D)
            {
                TaskDialog.Show("HMV Tools - Error",
                    "Cannot place annotations in a 3D view.\n" +
                    "Please switch to a plan, section, or detail view.");
                return Result.Failed;
            }

            // ─ 2. Select Pipes / FlexPipes ─────────────────────────
            IList<Element> selectedElements;
            try
            {
                selectedElements = SelectPipesAndFlexPipes(uidoc);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }

            if (selectedElements == null || selectedElements.Count == 0)
            {
                TaskDialog.Show("HMV Tools",
                    "No Pipe or FlexPipe elements selected.");
                return Result.Cancelled;
            }

            // ─ 3. Collect families for both categories ─────────────
            List<FamilyEntry> annotEntries =
                CollectFamilyEntries(doc,
                    BuiltInCategory.OST_GenericAnnotation);

            List<FamilyEntry> detailEntries =
                CollectFamilyEntries(doc,
                    BuiltInCategory.OST_DetailComponents);

            if (annotEntries.Count == 0 && detailEntries.Count == 0)
            {
                TaskDialog.Show("HMV Tools - Error",
                    "No Generic Annotation or Detail Item families " +
                    "loaded in the project.");
                return Result.Failed;
            }

            // ─ 4. Show the selection window ────────────────────────
            var window = new PipeAnnotationWindow(
                annotEntries, detailEntries);
            bool? result = window.ShowDialog();

            if (result != true || window.Settings == null)
                return Result.Cancelled;

            PipeAnnotationSettings settings = window.Settings;

            // ─ 5. Resolve the chosen FamilySymbol ──────────────────
            BuiltInCategory targetCat =
                settings.Mode == PlacementMode.GenericAnnotation
                    ? BuiltInCategory.OST_GenericAnnotation
                    : BuiltInCategory.OST_DetailComponents;

            FamilySymbol symbol = FindSymbol(
                doc, targetCat, settings.FamilyName, settings.TypeName);

            if (symbol == null)
            {
                TaskDialog.Show("HMV Tools - Error",
                    $"Could not resolve family type:\n" +
                    $"{settings.FamilyName} : {settings.TypeName}");
                return Result.Failed;
            }

            // ─ 6. Execute based on mode ────────────────────────────
            if (settings.Mode == PlacementMode.GenericAnnotation)
                return ExecuteAnnotationMode(
                    doc, activeView, symbol, selectedElements, settings);
            else
                return ExecuteDetailItemMode(
                    doc, activeView, symbol, selectedElements);
        }

        // ────────────────────────────────────────────────────────────
        // MODE A: Generic Annotation (spaced along path)
        // ────────────────────────────────────────────────────────────

        private Result ExecuteAnnotationMode(
            Document doc,
            View view,
            FamilySymbol symbol,
            IList<Element> selectedElements,
            PipeAnnotationSettings settings)
        {
            double spacingFeet = MmToFeet(settings.SpacingMm);
            if (spacingFeet <= 1e-9)
            {
                TaskDialog.Show("HMV Tools - Error",
                    "Spacing must be greater than zero.");
                return Result.Failed;
            }

            int totalPlaced = 0;

            using (Transaction tx = new Transaction(doc,
                "Place Annotations Along Pipes"))
            {
                tx.Start();

                if (!symbol.IsActive)
                    symbol.Activate();

                foreach (Element elem in selectedElements)
                {
                    IList<Curve> curves = GetElementCurves(elem);
                    if (curves == null || curves.Count == 0)
                        continue;

                    foreach (Curve curve in curves)
                    {
                        totalPlaced += PlaceAnnotationsAlongCurve(
                            doc, view, symbol, curve, spacingFeet);
                    }
                }

                tx.Commit();
            }

            TaskDialog.Show("HMV Tools",
                $"Placed {totalPlaced} annotation(s) on " +
                $"{selectedElements.Count} element(s).");

            return Result.Succeeded;
        }

        // ────────────────────────────────────────────────────────────
        // MODE B: Detail Item (one instance start → end, linear only)
        // ────────────────────────────────────────────────────────────

        private Result ExecuteDetailItemMode(
            Document doc,
            View view,
            FamilySymbol symbol,
            IList<Element> selectedElements)
        {
            int totalPlaced = 0;
            int skippedCurved = 0;

            using (Transaction tx = new Transaction(doc,
                "Place Detail Items Along Pipes"))
            {
                tx.Start();

                if (!symbol.IsActive)
                    symbol.Activate();

                // ── Discover the EXACT Z of the view plane ─────
                // Create a temp detail line — Revit auto-projects
                // it onto the view's internal plane. Read back
                // the real Z, then delete the temp line.
                DetailCurve tempDC = doc.Create.NewDetailCurve(
                    view,
                    Line.CreateBound(
                        new XYZ(0, 0, 0),
                        new XYZ(1, 0, 0)));

                Line tempLine = (Line)((LocationCurve)
                    tempDC.Location).Curve;
                double viewZ = tempLine.GetEndPoint(0).Z;
                doc.Delete(tempDC.Id);

                // ── Place detail items ─────────────────────────
                foreach (Element elem in selectedElements)
                {
                    IList<Curve> curves = GetElementCurves(elem);
                    if (curves == null || curves.Count == 0)
                        continue;

                    foreach (Curve curve in curves)
                    {
                        if (!(curve is Line line))
                        {
                            skippedCurved++;
                            continue;
                        }

                        XYZ ptA = line.GetEndPoint(0);
                        XYZ ptB = line.GetEndPoint(1);

                        // Always orient from higher Z to lower Z
                        // (flow follows slope/gravity)
                        XYZ startPt = ptA.Z >= ptB.Z ? ptA : ptB;
                        XYZ endPt = ptA.Z >= ptB.Z ? ptB : ptA;

                        // Flatten to the EXACT view plane Z
                        XYZ projStart = new XYZ(
                            startPt.X, startPt.Y, viewZ);
                        XYZ projEnd = new XYZ(
                            endPt.X, endPt.Y, viewZ);

                        if (projStart.DistanceTo(projEnd) < 1e-6)
                        {
                            skippedCurved++;
                            continue;
                        }

                        Line placementLine =
                            Line.CreateBound(projStart, projEnd);

                        doc.Create.NewFamilyInstance(
                            placementLine, symbol, view);

                        totalPlaced++;
                    }
                }

                tx.Commit();
            }

            string msg = $"Placed {totalPlaced} detail item(s) on "
                + $"{selectedElements.Count} element(s).";

            if (skippedCurved > 0)
            {
                msg += $"\n\nSkipped {skippedCurved} curved segment(s) "
                    + "(arcs/splines not supported in "
                    + "Detail Item mode).";
            }

            TaskDialog.Show("HMV Tools", msg);
            return Result.Succeeded;
        }

        // ── COLLECT FAMILY ENTRIES ──────────────────────────────────

        /// <summary>
        /// Gathers all FamilySymbol types for a given category,
        /// sorted alphabetically by family name then type name.
        /// </summary>
        private List<FamilyEntry> CollectFamilyEntries(
            Document doc, BuiltInCategory category)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(category)
                .Cast<FamilySymbol>()
                .Select(fs => new FamilyEntry
                {
                    FamilyName = fs.FamilyName,
                    TypeName = fs.Name
                })
                .OrderBy(e => e.FamilyName)
                .ThenBy(e => e.TypeName)
                .ToList();
        }

        // ── SELECTION ──────────────────────────────────────────────

        private IList<Element> SelectPipesAndFlexPipes(UIDocument uidoc)
        {
            ICollection<ElementId> preSelected =
                uidoc.Selection.GetElementIds();

            if (preSelected.Count > 0)
            {
                Document doc = uidoc.Document;
                List<Element> filtered = preSelected
                    .Select(id => doc.GetElement(id))
                    .Where(e => e is Pipe
                        || e is FlexPipe
                        || e.Category?.Id.IntegerValue
                            == (int)BuiltInCategory
                                .OST_StructuralFraming)
                    .ToList();

                if (filtered.Count > 0)
                    return filtered;
            }

            IList<Reference> refs = uidoc.Selection.PickObjects(
                ObjectType.Element,
                new PipeFlexPipeFilter(),
                "Select Pipes, Flex Pipes, and/or Structural Framing, then press Finish."); 

            return refs
                .Select(r => uidoc.Document.GetElement(r))
                .ToList();
        }

        private class PipeFlexPipeFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                return elem is Pipe
                    || elem is FlexPipe
                    || elem.Category?.Id.IntegerValue
                        == (int)BuiltInCategory
                            .OST_StructuralFraming;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }

        // ── SYMBOL RESOLUTION ──────────────────────────────────────

        private FamilySymbol FindSymbol(
            Document doc, BuiltInCategory category,
            string familyName, string typeName)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(category)
                .Cast<FamilySymbol>()
                .FirstOrDefault(fs =>
                    fs.FamilyName.Equals(familyName,
                        StringComparison.OrdinalIgnoreCase)
                    && fs.Name.Equals(typeName,
                        StringComparison.OrdinalIgnoreCase));
        }

        // ── CURVE EXTRACTION ───────────────────────────────────────

        private IList<Curve> GetElementCurves(Element elem)
        {
            List<Curve> result = new List<Curve>();

            LocationCurve locCurve = elem.Location as LocationCurve;
            if (locCurve != null && locCurve.Curve != null)
            {
                result.Add(locCurve.Curve);
                return result;
            }

            Options geoOptions = new Options
            {
                ComputeReferences = false,
                DetailLevel = ViewDetailLevel.Fine,
                IncludeNonVisibleObjects = true
            };

            GeometryElement geoElem = elem.get_Geometry(geoOptions);
            if (geoElem != null)
                ExtractCurvesRecursive(geoElem, result);

            return result;
        }

        private void ExtractCurvesRecursive(
            GeometryElement geoElem, List<Curve> curves)
        {
            foreach (GeometryObject geoObj in geoElem)
            {
                if (geoObj is Curve crv)
                {
                    curves.Add(crv);
                }
                else if (geoObj is GeometryInstance gi)
                {
                    GeometryElement instGeo = gi.GetInstanceGeometry();
                    if (instGeo != null)
                        ExtractCurvesRecursive(instGeo, curves);
                }
            }
        }

        // ── ANNOTATION PLACEMENT (MODE A) ──────────────────────────

        private int PlaceAnnotationsAlongCurve(
            Document doc,
            View view,
            FamilySymbol symbol,
            Curve curve,
            double spacing)
        {
            double curveLength = curve.Length;
            if (curveLength < 1e-9)
                return 0;

            // Ensure curve flows from higher Z to lower Z
            // (gravity/slope direction). If reversed, flip it.
            XYZ csStart = curve.GetEndPoint(0);
            XYZ csEnd = curve.GetEndPoint(1);
            if (csStart.Z < csEnd.Z)
                curve = curve.CreateReversed();

            int count = (int)Math.Floor(curveLength / spacing);
            if (count < 1) count = 1;

            double totalOccupied = count * spacing;
            double startOffset =
                (curveLength - totalOccupied) / 2.0 + spacing / 2.0;

            XYZ viewNormal = view.ViewDirection;
            int placed = 0;

            for (int i = 0; i < count; i++)
            {
                double distAlongCurve = startOffset + i * spacing;

                if (distAlongCurve < 0 || distAlongCurve > curveLength)
                    continue;

                double normalizedParam = distAlongCurve / curveLength;
                double rawParam =
                    curve.ComputeRawParameter(normalizedParam);

                Transform derivatives =
                    curve.ComputeDerivatives(rawParam, false);

                XYZ point = derivatives.Origin;
                XYZ tangent = derivatives.BasisX.Normalize();

                FamilyInstance inst = doc.Create.NewFamilyInstance(
                    point, symbol, view);

                double angle = ComputeRotationAngle(tangent, view);

                if (Math.Abs(angle) > 1e-9)
                {
                    Line rotAxis = Line.CreateBound(
                        point, point + viewNormal);

                    ElementTransformUtils.RotateElement(
                        doc, inst.Id, rotAxis, angle);
                }

                placed++;
            }

            return placed;
        }

        // ── ROTATION ANGLE ─────────────────────────────────────────

        private double ComputeRotationAngle(XYZ tangent, View view)
        {
            double tx = tangent.DotProduct(view.RightDirection);
            double ty = tangent.DotProduct(view.UpDirection);
            return Math.Atan2(ty, tx);
        }

        /// <summary>
        /// Returns the Z elevation of the view's working plane.
        /// For plan views: level + cut plane offset.
        /// For other views: falls back to view origin Z.
        /// </summary>
        private double GetViewPlaneZ(View view)
        {
            if (view is ViewPlan vp && vp.GenLevel != null)
            {
                PlanViewRange pvr = vp.GetViewRange();
                double cutOffset = pvr.GetOffset(
                    PlanViewPlane.CutPlane);
                return vp.GenLevel.ProjectElevation + cutOffset;
            }

            return view.Origin.Z;
        }
    }
}