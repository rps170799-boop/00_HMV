using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class TopographyToLinesCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                FilteredElementCollector linkCollector =
                    new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance));

                List<RevitLinkInstance> links = new List<RevitLinkInstance>();
                List<string> linkNames = new List<string>();

                foreach (RevitLinkInstance link in linkCollector)
                {
                    links.Add(link);
                    linkNames.Add(link.Name);
                }

                if (links.Count == 0)
                {
                    TaskDialog.Show("HMV Tools", "No se encontraron vinculos de Revit en el documento.");
                    return Result.Failed;
                }

                TopographyToLinesWindow win1 = new TopographyToLinesWindow(linkNames);
                if (win1.ShowDialog() != true) return Result.Cancelled;

                int selectedIndex = win1.SelectedLinkIndex;
                if (selectedIndex < 0 || selectedIndex >= links.Count) return Result.Cancelled;

                RevitLinkInstance selectedLink = links[selectedIndex];
                bool generateOffsets = win1.GenerateOffsets;

                int created = 0, skipped = 0;

                if (generateOffsets)
                {
                    List<string> styleNames = new List<string>();
                    Category linesCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
                    if (linesCat != null)
                    {
                        foreach (Category subCat in linesCat.SubCategories) styleNames.Add(subCat.Name);
                    }
                    styleNames.Sort();

                    List<string> textStyleNames = new List<string>();
                    FilteredElementCollector textCat = new FilteredElementCollector(doc).OfClass(typeof(TextNoteType));
                    foreach (TextNoteType tnt in textCat) textStyleNames.Add(tnt.Name);
                    textStyleNames.Sort();

                    // GET DIMENSION STYLES
                    List<string> dimStyleNames = new List<string>();
                    FilteredElementCollector dimCat = new FilteredElementCollector(doc).OfClass(typeof(DimensionType));
                    foreach (DimensionType dt in dimCat)
                    {
                        if (dt.StyleType == DimensionStyleType.Linear)
                            dimStyleNames.Add(dt.Name);
                    }
                    dimStyleNames.Sort();

                    TopographyOffsetsWindow win2 = new TopographyOffsetsWindow(styleNames, textStyleNames, dimStyleNames);
                    if (win2.ShowDialog() != true) return Result.Cancelled;

                    ProcessTopography(doc, selectedLink, true, win2.Offset1, win2.Offset2, win2.Offset3, win2.SelectedLineStyle, win2.SelectedTextStyle, win2.SelectedDimensionStyle, out created, out skipped);
                }
                else
                {
                    ProcessTopography(doc, selectedLink, false, 0, 0, 0, null, null, null, out created, out skipped);
                }

                TaskDialog.Show("HMV Tools", $"Proceso completado:\n\nLineas creadas: {created}\nLineas omitidas: {skipped}");
                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", ex.Message);
                return Result.Failed;
            }
        }

        private void ProcessTopography(
            Document doc,
            RevitLinkInstance link,
            bool generateOffsets,
            double userOffset1_m,
            double userOffset2_m,
            double userOffset3_m,
            string userStyleName,
            string userTextStyleName,
            string userDimStyleName, // NEW
            out int created,
            out int skipped)
        {
            created = 0;
            skipped = 0;

            View activeView = doc.ActiveView;
            XYZ origin = activeView.Origin;
            XYZ viewDir = activeView.ViewDirection;

            Document linkDoc = link.GetLinkDocument();
            if (linkDoc == null) return;

            Transform transform = link.GetTotalTransform();
            List<Mesh> meshes = new List<Mesh>();
            FilteredElementCollector topoCollector = new FilteredElementCollector(linkDoc)
                .OfCategory(BuiltInCategory.OST_Topography)
                .WhereElementIsNotElementType();

            Options opt = new Options();
            foreach (Element topo in topoCollector)
            {
                GeometryElement geom = topo.get_Geometry(opt);
                if (geom == null) continue;
                foreach (GeometryObject gObj in geom)
                {
                    if (gObj is Mesh mesh) meshes.Add(mesh);
                }
            }

            List<Tuple<XYZ, XYZ>> rawSegments = ExtractIntersectionSegments(meshes, transform, origin, viewDir);
            List<Tuple<XYZ, XYZ>> segments = MergeCollinearSegments(rawSegments);

            // CLIPPING LOGIC FOR TEXT MARGIN
            BoundingBoxXYZ cropBox = activeView.CropBox;
            Transform bTransform = cropBox != null ? cropBox.Transform : Transform.Identity;
            double minXLocal = double.MaxValue;
            XYZ extremeLeftPoint = null;

            if (cropBox != null)
            {
                foreach (var seg in segments)
                {
                    XYZ p1Local = bTransform.Inverse.OfPoint(seg.Item1);
                    XYZ p2Local = bTransform.Inverse.OfPoint(seg.Item2);

                    if (p1Local.X < minXLocal) { minXLocal = p1Local.X; extremeLeftPoint = seg.Item1; }
                    if (p2Local.X < minXLocal) { minXLocal = p2Local.X; extremeLeftPoint = seg.Item2; }
                }
            }

            List<Tuple<XYZ, XYZ>> clippedSegments = new List<Tuple<XYZ, XYZ>>();
            double textMarginMeters = 3.5;
            double textMarginFeet = textMarginMeters * 3.28084;
            double clipXLocal = minXLocal + textMarginFeet;

            if (cropBox != null && generateOffsets)
            {
                foreach (var seg in segments)
                {
                    XYZ p1Local = bTransform.Inverse.OfPoint(seg.Item1);
                    XYZ p2Local = bTransform.Inverse.OfPoint(seg.Item2);

                    if (p1Local.X > p2Local.X) { XYZ temp = p1Local; p1Local = p2Local; p2Local = temp; }

                    if (p2Local.X <= clipXLocal) continue;

                    if (p1Local.X < clipXLocal)
                    {
                        double t = (clipXLocal - p1Local.X) / (p2Local.X - p1Local.X);
                        p1Local = p1Local + t * (p2Local - p1Local);
                    }
                    clippedSegments.Add(new Tuple<XYZ, XYZ>(bTransform.OfPoint(p1Local), bTransform.OfPoint(p2Local)));
                }
            }
            else { clippedSegments = segments; }

            // ── LIFTED OFFSET VARIABLES FOR TEXT CALCULATIONS ──
            Transform t1 = null, t2 = null, t3 = null, t4 = null;
            double dist1 = 0, dist2 = 0, dist3 = 0, dist4 = 0;

            if (generateOffsets)
            {
                double initialOffset_mm = 100.0;
                double mmToFeet = 1.0 / 304.8;

                dist1 = initialOffset_mm * mmToFeet;
                dist2 = dist1 + (userOffset1_m * mmToFeet);
                dist3 = dist2 + (userOffset2_m * mmToFeet);
                dist4 = dist1 + (userOffset3_m * mmToFeet);

                XYZ dir = activeView.UpDirection;
                t1 = Transform.CreateTranslation(dir.Multiply(dist1));
                t2 = Transform.CreateTranslation(dir.Multiply(dist2));
                t3 = Transform.CreateTranslation(dir.Multiply(dist3));
                t4 = Transform.CreateTranslation(dir.Multiply(dist4));
            }

            using (Transaction tx = new Transaction(doc, "HMV - Topography to Lines"))
            {
                tx.Start();

                GraphicsStyle thinStyle = GetLineStyleByName(doc, "<Thin Lines>");
                if (thinStyle == null) thinStyle = GetLineStyleByName(doc, "<Líneas finas>");

                GraphicsStyle userSelectedStyle = generateOffsets ? GetLineStyleByName(doc, userStyleName) : null;
                List<ElementId> linesToGroup = new List<ElementId>();

                // Variables for Dimensions
                DetailCurve leftDc1 = null;
                DetailCurve leftDc2 = null;
                DetailCurve leftDc3 = null;
                DetailCurve leftDc4 = null;

                foreach (Tuple<XYZ, XYZ> seg in clippedSegments)
                {
                    try
                    {
                        Line baseLine = Line.CreateBound(seg.Item1, seg.Item2);
                        DetailCurve baseCurve = doc.Create.NewDetailCurve(activeView, baseLine);
                        if (thinStyle != null) baseCurve.LineStyle = thinStyle;

                        linesToGroup.Add(baseCurve.Id);
                        created++;

                        if (generateOffsets)
                        {
                            Line line1 = baseLine.CreateTransformed(t1) as Line;
                            DetailCurve dc1 = doc.Create.NewDetailCurve(activeView, line1);
                            if (thinStyle != null) dc1.LineStyle = thinStyle;
                            linesToGroup.Add(dc1.Id);
                            created++;

                            Line line2 = baseLine.CreateTransformed(t2) as Line;
                            DetailCurve dc2 = doc.Create.NewDetailCurve(activeView, line2);
                            if (userSelectedStyle != null) dc2.LineStyle = userSelectedStyle;
                            linesToGroup.Add(dc2.Id);
                            created++;

                            Line line3 = baseLine.CreateTransformed(t3) as Line;
                            DetailCurve dc3 = doc.Create.NewDetailCurve(activeView, line3);
                            if (userSelectedStyle != null) dc3.LineStyle = userSelectedStyle;
                            linesToGroup.Add(dc3.Id);
                            created++;

                            Line line4 = baseLine.CreateTransformed(t4) as Line;
                            DetailCurve dc4 = doc.Create.NewDetailCurve(activeView, line4);
                            if (userSelectedStyle != null) dc4.LineStyle = userSelectedStyle;
                            linesToGroup.Add(dc4.Id);
                            created++;

                            // Save the very first/leftmost group of curves to use for dimensioning
                            if (leftDc1 == null && extremeLeftPoint != null &&
                               (seg.Item1.DistanceTo(extremeLeftPoint) < 1e-4 || seg.Item2.DistanceTo(extremeLeftPoint) < 1e-4))
                            {
                                leftDc1 = dc1;
                                leftDc2 = dc2;
                                leftDc3 = dc3;
                                leftDc4 = dc4;
                            }
                        }
                    }
                    catch { skipped++; }
                }

                // --- TEXT AND DIMENSION ANNOTATION ---
                if (generateOffsets && cropBox != null && extremeLeftPoint != null)
                {
                    try
                    {
                        // 1. First, find the Text Note Type 
                        TextNoteType selectedTextType = null;
                        FilteredElementCollector txtCollector = new FilteredElementCollector(doc).OfClass(typeof(TextNoteType));
                        foreach (TextNoteType t in txtCollector)
                        {
                            if (t.Name.Equals(userTextStyleName, StringComparison.OrdinalIgnoreCase))
                            {
                                selectedTextType = t as TextNoteType;
                                break;
                            }
                        }

                        // 2. Setup the text coordinates
                        XYZ basePointLocal = bTransform.Inverse.OfPoint(extremeLeftPoint);
                        double yMidBottom = basePointLocal.Y + (dist1 + dist2) / 2.0;
                        double yMidTop = basePointLocal.Y + (dist2 + dist3) / 2.0;
                        XYZ textPosBottomWorld = bTransform.OfPoint(new XYZ(cropBox.Min.X, yMidBottom, 0));
                        XYZ textPosTopWorld = bTransform.OfPoint(new XYZ(cropBox.Min.X, yMidTop, 0));

                        // 3. Setup Text Options
                        TextNoteOptions opts = new TextNoteOptions();
                        opts.HorizontalAlignment = HorizontalTextAlignment.Left;
                        opts.VerticalAlignment = VerticalTextAlignment.Middle;
                        opts.TypeId = selectedTextType != null ? selectedTextType.Id : doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);

                        // Lower Gap (Distancia)
                        TextNote note1 = TextNote.Create(doc, activeView.Id, textPosBottomWorld, "DISTANCIA DE\r\nSEGURIDAD", opts);
                        linesToGroup.Add(note1.Id);

                        // Upper Gap (Valor Basico)
                        TextNote note2 = TextNote.Create(doc, activeView.Id, textPosTopWorld, "VALOR BÁSICO", opts);
                        linesToGroup.Add(note2.Id);

                        // --- NEW: DIMENSIONS ---
                        if (leftDc1 != null && !string.IsNullOrEmpty(userDimStyleName))
                        {
                            DimensionType selectedDimType = null;
                            FilteredElementCollector dimCollector = new FilteredElementCollector(doc).OfClass(typeof(DimensionType));
                            foreach (DimensionType dt in dimCollector)
                            {
                                if (dt.Name.Equals(userDimStyleName, StringComparison.OrdinalIgnoreCase))
                                {
                                    selectedDimType = dt as DimensionType;
                                    break;
                                }
                            }

                            if (selectedDimType != null)
                            {
                                // Create vertical line to place dimensions on
                                double dimX = cropBox.Min.X + (1.5 * 3.28084); // Place slightly to the right of text
                                XYZ p1 = bTransform.OfPoint(new XYZ(dimX, -100, 0));
                                XYZ p2 = bTransform.OfPoint(new XYZ(dimX, 100, 0));
                                Line dimLine = Line.CreateBound(p1, p2);

                                Action<DetailCurve, DetailCurve> CreateDim = (c1, c2) =>
                                {
                                    ReferenceArray ra = new ReferenceArray();
                                    ra.Append(c1.GeometryCurve.Reference);
                                    ra.Append(c2.GeometryCurve.Reference);
                                    Dimension dim = doc.Create.NewDimension(activeView, dimLine, ra, selectedDimType);
                                    if (dim != null) linesToGroup.Add(dim.Id);
                                };

                                // Create the requested dimensions
                                CreateDim(leftDc1, leftDc2); // 0.10 offset to 1st offset
                                CreateDim(leftDc2, leftDc3); // 1st to 2nd offset
                                CreateDim(leftDc1, leftDc4); // 0.10 offset to last offset
                            }
                        }
                    }
                    catch { }
                }

                if (linesToGroup.Count > 0)
                {
                    Group newGroup = doc.Create.NewGroup(linesToGroup);
                    string baseName = "HMV_LineasGrupo";
                    int suffix = 1;
                    bool nameSet = false;

                    while (!nameSet)
                    {
                        try { newGroup.GroupType.Name = baseName + suffix; nameSet = true; }
                        catch (Autodesk.Revit.Exceptions.ArgumentException) { suffix++; }
                    }
                }

                tx.Commit();
            }
        }

        private GraphicsStyle GetLineStyleByName(Document doc, string styleName)
        {
            if (string.IsNullOrEmpty(styleName)) return null;
            Category linesCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
            if (linesCat != null)
            {
                foreach (Category subCat in linesCat.SubCategories)
                {
                    if (subCat.Name.Equals(styleName, StringComparison.OrdinalIgnoreCase))
                        return subCat.GetGraphicsStyle(GraphicsStyleType.Projection);
                }
            }
            return null;
        }

        private List<Tuple<XYZ, XYZ>> ExtractIntersectionSegments(List<Mesh> meshes, Transform transform, XYZ origin, XYZ normal)
        {
            List<Tuple<XYZ, XYZ>> segments = new List<Tuple<XYZ, XYZ>>();
            foreach (Mesh mesh in meshes)
            {
                for (int i = 0; i < mesh.NumTriangles; i++)
                {
                    MeshTriangle tri = mesh.get_Triangle(i);
                    XYZ[] verts = new XYZ[3];
                    for (int j = 0; j < 3; j++) verts[j] = transform.OfPoint(tri.get_Vertex(j));

                    double[] dists = new double[3];
                    for (int j = 0; j < 3; j++) dists[j] = SignedDistance(verts[j], origin, normal);

                    List<XYZ> intPts = new List<XYZ>();
                    int[,] edges = { { 0, 1 }, { 1, 2 }, { 2, 0 } };

                    for (int e = 0; e < 3; e++)
                    {
                        int e0 = edges[e, 0];
                        int e1 = edges[e, 1];
                        double d0 = dists[e0];
                        double d1 = dists[e1];

                        if (d0 * d1 < 0) intPts.Add(Interpolate(verts[e0], verts[e1], d0, d1));
                        else if (Math.Abs(d0) < 1e-9) intPts.Add(verts[e0]);
                    }

                    List<XYZ> unique = new List<XYZ>();
                    foreach (XYZ p in intPts)
                    {
                        bool isDup = false;
                        foreach (XYZ u in unique) { if (p.DistanceTo(u) < 1e-6) { isDup = true; break; } }
                        if (!isDup) unique.Add(p);
                    }
                    if (unique.Count == 2 && unique[0].DistanceTo(unique[1]) > 1e-6) segments.Add(Tuple.Create(unique[0], unique[1]));
                }
            }
            return segments;
        }

        private double SignedDistance(XYZ point, XYZ origin, XYZ normal) { return (point - origin).DotProduct(normal); }

        private XYZ Interpolate(XYZ v0, XYZ v1, double d0, double d1)
        {
            double t = d0 / (d0 - d1);
            return new XYZ(v0.X + t * (v1.X - v0.X), v0.Y + t * (v1.Y - v0.Y), v0.Z + t * (v1.Z - v0.Z));
        }

        private List<Tuple<XYZ, XYZ>> MergeCollinearSegments(List<Tuple<XYZ, XYZ>> inputSegments)
        {
            if (inputSegments == null || inputSegments.Count < 2) return inputSegments;
            List<Tuple<XYZ, XYZ>> result = new List<Tuple<XYZ, XYZ>>();
            List<Tuple<XYZ, XYZ>> pool = new List<Tuple<XYZ, XYZ>>(inputSegments);
            double tol = 1e-4;

            while (pool.Count > 0)
            {
                var current = pool[0];
                pool.RemoveAt(0);
                XYZ start = current.Item1;
                XYZ end = current.Item2;
                bool extended = true;

                while (extended)
                {
                    extended = false;
                    for (int i = 0; i < pool.Count; i++)
                    {
                        var candidate = pool[i];
                        if (candidate.Item1.DistanceTo(end) < tol)
                        {
                            if (AreCollinear(start, end, candidate.Item2)) { end = candidate.Item2; pool.RemoveAt(i); extended = true; break; }
                        }
                        else if (candidate.Item2.DistanceTo(end) < tol)
                        {
                            if (AreCollinear(start, end, candidate.Item1)) { end = candidate.Item1; pool.RemoveAt(i); extended = true; break; }
                        }
                        else if (candidate.Item2.DistanceTo(start) < tol)
                        {
                            if (AreCollinear(candidate.Item1, start, end)) { start = candidate.Item1; pool.RemoveAt(i); extended = true; break; }
                        }
                        else if (candidate.Item1.DistanceTo(start) < tol)
                        {
                            if (AreCollinear(candidate.Item2, start, end)) { start = candidate.Item2; pool.RemoveAt(i); extended = true; break; }
                        }
                    }
                }
                if (start.DistanceTo(end) > 0.005) result.Add(new Tuple<XYZ, XYZ>(start, end));
            }
            return result;
        }

        private bool AreCollinear(XYZ p1, XYZ p2, XYZ p3)
        {
            if (p1.DistanceTo(p2) < 1e-5 || p2.DistanceTo(p3) < 1e-5) return true;
            XYZ v1 = (p2 - p1).Normalize();
            XYZ v2 = (p3 - p2).Normalize();
            return v1.DotProduct(v2) > 0.99999;
        }
    }
}