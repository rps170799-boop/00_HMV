using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class TopographyToLinesCommand : IExternalCommand
    {
        private class OffsetSettings
        {
            public double Offset1_mm;   // DIST. SEGURIDAD (from 0.10m line)
            public double Offset2_mm;   // VALOR BASICO    (from dist. seg. line)
            public double Offset3_mm;   // NIVEL CONEXION  (from 0.10m line)
            public string LineStyleName;
            public string DimStyleName;
            public string TextStyleName;
        }

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // ═══════════════════════════════════════════════════════
                // 1. COLLECT REVIT LINKS
                // ═══════════════════════════════════════════════════════
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
                    TaskDialog.Show("HMV Tools",
                        "No se encontraron vínculos de Revit en el documento.");
                    return Result.Failed;
                }

                // ═══════════════════════════════════════════════════════
                // 2. WINDOW 1 — Select link + checkbox
                // ═══════════════════════════════════════════════════════
                TopographyToLinesWindow win1 = new TopographyToLinesWindow(linkNames);
                if (win1.ShowDialog() != true) return Result.Cancelled;

                int selectedIndex = win1.SelectedLinkIndex;
                if (selectedIndex < 0 || selectedIndex >= links.Count) return Result.Cancelled;

                RevitLinkInstance selectedLink = links[selectedIndex];
                bool generateOffsets = win1.GenerateOffsets;

                // ═══════════════════════════════════════════════════════
                // 3A. CHECKBOX OFF → base topo line only, active view
                // ═══════════════════════════════════════════════════════
                if (!generateOffsets)
                {
                    int created = 0, skipped = 0;
                    ProcessSingleView(doc, selectedLink, doc.ActiveView,
                        false, null, 1, out created, out skipped);

                    TaskDialog.Show("HMV Tools",
                        $"Proceso completado:\n\n" +
                        $"Líneas creadas: {created}\n" +
                        $"Líneas omitidas: {skipped}");
                    return Result.Succeeded;
                }

                // ═══════════════════════════════════════════════════════
                // 3B. CHECKBOX ON → Window 2 (styles + offsets)
                // ═══════════════════════════════════════════════════════
                List<string> styleNames = new List<string>();
                Category linesCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
                if (linesCat != null)
                {
                    foreach (Category subCat in linesCat.SubCategories)
                        styleNames.Add(subCat.Name);
                }
                styleNames.Sort();

                List<string> textStyleNames = new List<string>();
                foreach (TextNoteType tnt in new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNoteType)))
                    textStyleNames.Add(tnt.Name);
                textStyleNames.Sort();

                // LINEAR DIMENSION TYPES ONLY
                List<string> dimStyleNames = new List<string>();
                foreach (DimensionType dt in new FilteredElementCollector(doc)
                    .OfClass(typeof(DimensionType)))
                {
                    if (dt.StyleType == DimensionStyleType.Linear)
                        dimStyleNames.Add(dt.Name);
                }
                dimStyleNames.Sort();

                TopographyOffsetsWindow win2 = new TopographyOffsetsWindow(
                    styleNames, textStyleNames, dimStyleNames);
                if (win2.ShowDialog() != true) return Result.Cancelled;

                OffsetSettings settings = new OffsetSettings
                {
                    Offset1_mm = win2.Offset1,
                    Offset2_mm = win2.Offset2,
                    Offset3_mm = win2.Offset3,
                    LineStyleName = win2.SelectedLineStyle,
                    DimStyleName = win2.SelectedDimensionStyle,
                    TextStyleName = win2.SelectedTextStyle
                };

                // ═══════════════════════════════════════════════════════
                // 4. WINDOW 3 — Select section views
                // ═══════════════════════════════════════════════════════
                var sectionViews = new List<View>();
                var sectionItems = new List<SectionViewItem>();
                ElementId activeId = doc.ActiveView.Id;

                var viewCollector = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSection))
                    .Cast<ViewSection>()
                    .Where(v => !v.IsTemplate)
                    .OrderBy(v => v.Name)
                    .ToList();

                bool activeAdded = false;

                for (int i = 0; i < viewCollector.Count; i++)
                {
                    sectionViews.Add(viewCollector[i]);
                    bool isActive = viewCollector[i].Id == activeId;
                    if (isActive) activeAdded = true;

                    sectionItems.Add(new SectionViewItem
                    {
                        Name = viewCollector[i].Name,
                        OriginalIndex = i,
                        IsActiveView = isActive
                    });
                }

                // If active view is not a section, insert at position 0
                if (!activeAdded)
                {
                    sectionViews.Insert(0, doc.ActiveView);
                    foreach (var si in sectionItems) si.OriginalIndex += 1;

                    sectionItems.Insert(0, new SectionViewItem
                    {
                        Name = doc.ActiveView.Name,
                        OriginalIndex = 0,
                        IsActiveView = true
                    });
                }

                if (sectionItems.Count == 0)
                {
                    TaskDialog.Show("HMV Tools",
                        "No se encontraron vistas de sección.");
                    return Result.Failed;
                }

                var win3 = new TopographySectionWindow(sectionItems);
                if (win3.ShowDialog() != true) return Result.Cancelled;
                if (win3.SelectedViewIndices.Count == 0) return Result.Cancelled;

                // ═══════════════════════════════════════════════════════
                // 5. PROCESS EACH SELECTED VIEW
                // ═══════════════════════════════════════════════════════
                int totalCreated = 0, totalSkipped = 0, groupCount = 0;
                List<string> viewResults = new List<string>();

                foreach (int viewIdx in win3.SelectedViewIndices)
                {
                    if (viewIdx < 0 || viewIdx >= sectionViews.Count) continue;

                    View targetView = sectionViews[viewIdx];
                    int created = 0, skipped = 0;
                    groupCount++;

                    ProcessSingleView(doc, selectedLink, targetView,
                        true, settings, groupCount, out created, out skipped);

                    totalCreated += created;
                    totalSkipped += skipped;
                    viewResults.Add($"  • {targetView.Name}: {created} líneas");
                }

                // ═══════════════════════════════════════════════════════
                // 6. REPORT
                // ═══════════════════════════════════════════════════════
                string report =
                    $"Proceso completado:\n\n" +
                    $"Vistas procesadas: {groupCount}\n" +
                    $"Total líneas creadas: {totalCreated}\n" +
                    $"Total líneas omitidas: {totalSkipped}\n\n" +
                    string.Join("\n", viewResults);

                TaskDialog.Show("HMV Tools", report);
                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", ex.ToString());
                return Result.Failed;
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // CORE: Process a single view
        //
        // Key fix for non-active views: after computing intersection
        // points, project them onto the EXACT view plane to eliminate
        // floating-point drift. Without this, NewDetailCurve fails or
        // produces flat lines in sections other than the active view.
        // ═══════════════════════════════════════════════════════════════
        // ═══════════════════════════════════════════════════════════════
        // CORE: Process a single view
        // ═══════════════════════════════════════════════════════════════
        // ═══════════════════════════════════════════════════════════════
        // CORE: Process a single view
        // ═══════════════════════════════════════════════════════════════
        private void ProcessSingleView(
            Document doc,
            RevitLinkInstance link,
            View targetView,
            bool generateOffsets,
            OffsetSettings settings,
            int groupNumber,
            out int created,
            out int skipped)
        {
            created = 0;
            skipped = 0;

            Document linkDoc = link.GetLinkDocument();
            if (linkDoc == null) return;

            Transform linkTransform = link.GetTotalTransform();

            // ── Extract meshes from linked topo ──
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

            if (meshes.Count == 0) return;

            // ── Use TARGET VIEW's plane for intersection ──
            XYZ origin = targetView.Origin;
            XYZ viewDir = targetView.ViewDirection;

            List<Tuple<XYZ, XYZ>> rawSegments =
                ExtractIntersectionSegments(meshes, linkTransform, origin, viewDir);

            // ── CRITICAL: Project every point onto the exact view plane ──
            List<Tuple<XYZ, XYZ>> projectedSegments = new List<Tuple<XYZ, XYZ>>();
            foreach (var seg in rawSegments)
            {
                XYZ p1 = ProjectOntoPlane(seg.Item1, origin, viewDir);
                XYZ p2 = ProjectOntoPlane(seg.Item2, origin, viewDir);
                if (p1.DistanceTo(p2) > 1e-6)
                    projectedSegments.Add(Tuple.Create(p1, p2));
            }

            List<Tuple<XYZ, XYZ>> segments = MergeCollinearSegments(projectedSegments);

            if (segments.Count == 0) return;

            // ── Crop box data ──
            BoundingBoxXYZ cropBox = targetView.CropBox;
            Transform bTransform = cropBox != null ? cropBox.Transform : Transform.Identity;

            // ── STRICT CLIPPING TO CROP BOX (ZERO OFFSET) ──
            List<Tuple<XYZ, XYZ>> clippedSegments = new List<Tuple<XYZ, XYZ>>();

            double viewLeftX = cropBox != null ? cropBox.Min.X : 0;
            double viewRightX = cropBox != null ? cropBox.Max.X : 0;

            double leftElevation = 0;
            bool foundElevation = false;
            double minFoundX = double.MaxValue;

            if (cropBox != null)
            {
                foreach (var seg in segments)
                {
                    XYZ p1Local = bTransform.Inverse.OfPoint(seg.Item1);
                    XYZ p2Local = bTransform.Inverse.OfPoint(seg.Item2);

                    if (p1Local.X > p2Local.X)
                    {
                        XYZ temp = p1Local; p1Local = p2Local; p2Local = temp;
                    }

                    // Skip if completely outside the visible view
                    if (p2Local.X <= viewLeftX || p1Local.X >= viewRightX) continue;

                    // Clip exactly to left boundary (no gap)
                    if (p1Local.X < viewLeftX)
                    {
                        double t = (viewLeftX - p1Local.X) / (p2Local.X - p1Local.X);
                        p1Local = p1Local + t * (p2Local - p1Local);
                    }

                    // Clip exactly to right boundary
                    if (p2Local.X > viewRightX)
                    {
                        double t = (viewRightX - p1Local.X) / (p2Local.X - p1Local.X);
                        p2Local = p1Local + t * (p2Local - p1Local);
                    }

                    clippedSegments.Add(Tuple.Create(bTransform.OfPoint(p1Local), bTransform.OfPoint(p2Local)));

                    // Track elevation at the far left for text & dims
                    if (p1Local.X < minFoundX)
                    {
                        minFoundX = p1Local.X;
                        leftElevation = p1Local.Y;
                        foundElevation = true;
                    }
                }
            }
            else
            {
                clippedSegments = segments;
                if (segments.Count > 0)
                {
                    minFoundX = bTransform.Inverse.OfPoint(segments[0].Item1).X;
                    leftElevation = bTransform.Inverse.OfPoint(segments[0].Item1).Y;
                    foundElevation = true;
                }
            }

            if (clippedSegments.Count == 0) return;

            // ── Offset distances (all in feet) ──
            Transform t1 = null, t2 = null, t3 = null, t4 = null;
            double dist1 = 0, dist2 = 0, dist3 = 0, dist4 = 0;

            if (generateOffsets)
            {
                double mmToFeet = 1.0 / 304.8;
                double initialOffset_mm = 100.0;

                dist1 = initialOffset_mm * mmToFeet;
                dist2 = dist1 + (settings.Offset1_mm * mmToFeet);
                dist3 = dist2 + (settings.Offset2_mm * mmToFeet);
                dist4 = dist1 + (settings.Offset3_mm * mmToFeet);

                XYZ dir = targetView.UpDirection;
                t1 = Transform.CreateTranslation(dir.Multiply(dist1));
                t2 = Transform.CreateTranslation(dir.Multiply(dist2));
                t3 = Transform.CreateTranslation(dir.Multiply(dist3));
                t4 = Transform.CreateTranslation(dir.Multiply(dist4));
            }

            // ── Line styles ──
            GraphicsStyle thinStyle = GetLineStyleByName(doc, "<Thin Lines>");
            if (thinStyle == null) thinStyle = GetLineStyleByName(doc, "<Líneas finas>");
            if (thinStyle == null) thinStyle = GetLineStyleByName(doc, "<Lineas finas>");

            GraphicsStyle userStyle = generateOffsets
                ? GetLineStyleByName(doc, settings.LineStyleName) : null;

            using (Transaction tx = new Transaction(doc, "HMV - Topography to Lines"))
            {
                tx.Start();

                List<ElementId> groupIds = new List<ElementId>();

                // Track leftmost segment's offset curves for dimensioning
                double minXForDim = double.MaxValue;
                DetailCurve leftDc1 = null;
                DetailCurve leftDc2 = null;
                DetailCurve leftDc3 = null;
                DetailCurve leftDc4 = null;

                foreach (var seg in clippedSegments)
                {
                    try
                    {
                        Line baseLine = Line.CreateBound(seg.Item1, seg.Item2);

                        // ── BASE TOPOGRAPHY LINE ──
                        DetailCurve baseCurve = doc.Create.NewDetailCurve(targetView, baseLine);
                        if (thinStyle != null) baseCurve.LineStyle = thinStyle;
                        groupIds.Add(baseCurve.Id);
                        created++;

                        if (!generateOffsets) continue;

                        DetailCurve dc1 = doc.Create.NewDetailCurve(targetView, baseLine.CreateTransformed(t1) as Line);
                        if (thinStyle != null) dc1.LineStyle = thinStyle;
                        groupIds.Add(dc1.Id);
                        created++;

                        DetailCurve dc2 = doc.Create.NewDetailCurve(targetView, baseLine.CreateTransformed(t2) as Line);
                        if (userStyle != null) dc2.LineStyle = userStyle;
                        groupIds.Add(dc2.Id);
                        created++;

                        DetailCurve dc3 = doc.Create.NewDetailCurve(targetView, baseLine.CreateTransformed(t3) as Line);
                        if (userStyle != null) dc3.LineStyle = userStyle;
                        groupIds.Add(dc3.Id);
                        created++;

                        DetailCurve dc4 = doc.Create.NewDetailCurve(targetView, baseLine.CreateTransformed(t4) as Line);
                        if (userStyle != null) dc4.LineStyle = userStyle;
                        groupIds.Add(dc4.Id);
                        created++;

                        double currentLocalX = bTransform.Inverse.OfPoint(seg.Item1).X;
                        if (currentLocalX <= minXForDim + 1e-5)
                        {
                            minXForDim = currentLocalX;
                            leftDc1 = dc1;
                            leftDc2 = dc2;
                            leftDc3 = dc3;
                            leftDc4 = dc4;
                        }
                    }
                    catch { skipped++; }
                }

                // ── TEXT AND DIMENSIONS ──
                if (generateOffsets && cropBox != null && foundElevation)
                {
                    try
                    {
                        TextNoteType selectedTextType = null;
                        foreach (TextNoteType tnt in new FilteredElementCollector(doc)
                            .OfClass(typeof(TextNoteType)))
                        {
                            if (tnt.Name.Equals(settings.TextStyleName, StringComparison.OrdinalIgnoreCase))
                            { selectedTextType = tnt; break; }
                        }

                        double yMidBottom = leftElevation + (dist1 + dist2) / 2.0;
                        double yMidTop = leftElevation + (dist2 + dist3) / 2.0;

                        // 1. GUARANTEE ZERO OFFSET: Use the exact crop box left boundary
                        double startX = cropBox.Min.X;

                        // 2. TEXT PLACEMENT: Pushed further to the right (4.0 meters from the edge)
                        // (Keeping this exactly as you had it, since it is OKAY)
                        double textX = startX + (2.0 * 3.28084);

                        XYZ textPosBottom = bTransform.OfPoint(new XYZ(textX, yMidBottom, 0));
                        XYZ textPosTop = bTransform.OfPoint(new XYZ(textX, yMidTop, 0));

                        TextNoteOptions opts = new TextNoteOptions();
                        opts.HorizontalAlignment = HorizontalTextAlignment.Left;
                        opts.VerticalAlignment = VerticalTextAlignment.Middle;
                        opts.TypeId = selectedTextType != null
                            ? selectedTextType.Id
                            : doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);

                        TextNote note1 = TextNote.Create(doc, targetView.Id,
                            textPosBottom, "DISTANCIA DE\r\nSEGURIDAD", opts);
                        groupIds.Add(note1.Id);

                        TextNote note2 = TextNote.Create(doc, targetView.Id,
                            textPosTop, "VALOR BÁSICO", opts);
                        groupIds.Add(note2.Id);

                        // ── DIMENSIONS ──
                        if (leftDc1 != null && !string.IsNullOrEmpty(settings.DimStyleName))
                        {
                            doc.Regenerate();

                            DimensionType selectedDimType = null;
                            foreach (DimensionType dt in new FilteredElementCollector(doc)
                                .OfClass(typeof(DimensionType)))
                            {
                                if (dt.StyleType == DimensionStyleType.Linear &&
                                    dt.Name.Equals(settings.DimStyleName, StringComparison.OrdinalIgnoreCase))
                                { selectedDimType = dt; break; }
                            }

                            if (selectedDimType != null)
                            {
                                // 3. DIMENSIONS ANCHORED TO TEXT & LEFT CORNER

                                // Inner dims (2000, 2310) placed right next to the text (0.5m left of the text)
                                double innerX = textX - (3.0 * 3.28084);
                                XYZ iP1 = bTransform.OfPoint(new XYZ(innerX, leftElevation - 100, 0));
                                XYZ iP2 = bTransform.OfPoint(new XYZ(innerX, leftElevation + 100, 0));
                                Line innerDimLine = Line.CreateBound(iP1, iP2);

                                // Outer dim (5000) dragged all the way left to exactly end at the corner 
                                // (0.15m padding so it's visible but right against the boundary)
                                double outerX = startX - (2.50 * 3.28084);
                                XYZ oP1 = bTransform.OfPoint(new XYZ(outerX, leftElevation - 100, 0));
                                XYZ oP2 = bTransform.OfPoint(new XYZ(outerX, leftElevation + 100, 0));
                                Line outerDimLine = Line.CreateBound(oP1, oP2);

                                CreateDim(doc, targetView, leftDc1, leftDc2, innerDimLine, selectedDimType, groupIds);
                                CreateDim(doc, targetView, leftDc2, leftDc3, innerDimLine, selectedDimType, groupIds);
                                CreateDim(doc, targetView, leftDc1, leftDc4, outerDimLine, selectedDimType, groupIds);
                            }
                        }
                    }
                    catch { /* annotation failed, continue with grouping */ }
                }
                // ── GROUP (unique name per view) ──
                if (groupIds.Count > 0)
                {
                    try
                    {
                        Group newGroup = doc.Create.NewGroup(groupIds);

                        string viewClean = targetView.Name
                            .Replace(" ", "_").Replace(":", "-").Replace("/", "-");
                        string desiredName = string.Format("HMV_Topo_{0:D2}_{1}",
                            groupNumber, viewClean);

                        bool nameSet = false;
                        string tryName = desiredName;
                        int retry = 0;
                        while (!nameSet && retry < 100)
                        {
                            try
                            {
                                newGroup.GroupType.Name = tryName;
                                nameSet = true;
                            }
                            catch
                            {
                                retry++;
                                tryName = desiredName + "_" + retry;
                            }
                        }
                    }
                    catch { }
                }

                tx.Commit();
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // DIMENSION HELPER
        // Uses the working pattern: doc.Regenerate() + GeometryCurve.Reference
        // ═══════════════════════════════════════════════════════════════
        private void CreateDim(
            Document doc, View view,
            DetailCurve c1, DetailCurve c2,
            Line dimLine, DimensionType dimType,
            List<ElementId> groupIds)
        {
            try
            {
                Reference ref1 = c1.GeometryCurve.Reference;
                Reference ref2 = c2.GeometryCurve.Reference;

                if (ref1 == null || ref2 == null) return;

                ReferenceArray ra = new ReferenceArray();
                ra.Append(ref1);
                ra.Append(ref2);

                Dimension dim = doc.Create.NewDimension(view, dimLine, ra, dimType);
                if (dim != null) groupIds.Add(dim.Id);
            }
            catch { }
        }

        // ═══════════════════════════════════════════════════════════════
        // PROJECT POINT ONTO PLANE
        // Eliminates floating-point drift (~1e-9) from intersection math.
        // Without this, NewDetailCurve produces flat/degenerate lines
        // when the target view is not the active view.
        // ═══════════════════════════════════════════════════════════════
        private XYZ ProjectOntoPlane(XYZ point, XYZ planeOrigin, XYZ planeNormal)
        {
            double signedDist = (point - planeOrigin).DotProduct(planeNormal);
            return point - planeNormal.Multiply(signedDist);
        }

        // ═══════════════════════════════════════════════════════════════
        // GEOMETRY: Triangle-plane intersection
        // ═══════════════════════════════════════════════════════════════
        private List<Tuple<XYZ, XYZ>> ExtractIntersectionSegments(
            List<Mesh> meshes, Transform transform, XYZ origin, XYZ normal)
        {
            List<Tuple<XYZ, XYZ>> segments = new List<Tuple<XYZ, XYZ>>();

            foreach (Mesh mesh in meshes)
            {
                for (int i = 0; i < mesh.NumTriangles; i++)
                {
                    MeshTriangle tri = mesh.get_Triangle(i);
                    XYZ[] verts = new XYZ[3];
                    for (int j = 0; j < 3; j++)
                        verts[j] = transform.OfPoint(tri.get_Vertex(j));

                    double[] dists = new double[3];
                    for (int j = 0; j < 3; j++)
                        dists[j] = SignedDistance(verts[j], origin, normal);

                    List<XYZ> intPts = new List<XYZ>();
                    int[,] edges = { { 0, 1 }, { 1, 2 }, { 2, 0 } };

                    for (int e = 0; e < 3; e++)
                    {
                        int e0 = edges[e, 0];
                        int e1 = edges[e, 1];
                        double d0 = dists[e0];
                        double d1 = dists[e1];

                        if (d0 * d1 < 0)
                            intPts.Add(Interpolate(verts[e0], verts[e1], d0, d1));
                        else if (Math.Abs(d0) < 1e-9)
                            intPts.Add(verts[e0]);
                    }

                    // Remove duplicates
                    List<XYZ> unique = new List<XYZ>();
                    foreach (XYZ p in intPts)
                    {
                        bool isDup = false;
                        foreach (XYZ u in unique)
                        {
                            if (p.DistanceTo(u) < 1e-6) { isDup = true; break; }
                        }
                        if (!isDup) unique.Add(p);
                    }

                    if (unique.Count == 2 && unique[0].DistanceTo(unique[1]) > 1e-6)
                        segments.Add(Tuple.Create(unique[0], unique[1]));
                }
            }
            return segments;
        }

        private double SignedDistance(XYZ point, XYZ origin, XYZ normal)
        {
            return (point - origin).DotProduct(normal);
        }

        private XYZ Interpolate(XYZ v0, XYZ v1, double d0, double d1)
        {
            double t = d0 / (d0 - d1);
            return new XYZ(
                v0.X + t * (v1.X - v0.X),
                v0.Y + t * (v1.Y - v0.Y),
                v0.Z + t * (v1.Z - v0.Z));
        }

        // ═══════════════════════════════════════════════════════════════
        // GEOMETRY: Merge collinear adjacent segments
        // Tolerance ~0.28% slope (tighter than 0.5%):
        //   cos(0.16°) ≈ 0.999996 — preserves topo detail
        // ═══════════════════════════════════════════════════════════════
        // Tolerance matches working Gemini value
        private const double COLLINEAR_DOT = 0.99999;

        private List<Tuple<XYZ, XYZ>> MergeCollinearSegments(
            List<Tuple<XYZ, XYZ>> inputSegments)
        {
            if (inputSegments == null || inputSegments.Count < 2)
                return inputSegments;

            List<Tuple<XYZ, XYZ>> result = new List<Tuple<XYZ, XYZ>>();
            List<Tuple<XYZ, XYZ>> pool = new List<Tuple<XYZ, XYZ>>(inputSegments);
            double endpointTol = 1e-4;

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
                        var cand = pool[i];

                        if (cand.Item1.DistanceTo(end) < endpointTol)
                        {
                            if (AreCollinear(start, end, cand.Item2))
                            { end = cand.Item2; pool.RemoveAt(i); extended = true; break; }
                        }
                        else if (cand.Item2.DistanceTo(end) < endpointTol)
                        {
                            if (AreCollinear(start, end, cand.Item1))
                            { end = cand.Item1; pool.RemoveAt(i); extended = true; break; }
                        }
                        else if (cand.Item2.DistanceTo(start) < endpointTol)
                        {
                            if (AreCollinear(cand.Item1, start, end))
                            { start = cand.Item1; pool.RemoveAt(i); extended = true; break; }
                        }
                        else if (cand.Item1.DistanceTo(start) < endpointTol)
                        {
                            if (AreCollinear(cand.Item2, start, end))
                            { start = cand.Item2; pool.RemoveAt(i); extended = true; break; }
                        }
                    }
                }

                if (start.DistanceTo(end) > 0.005)
                    result.Add(Tuple.Create(start, end));
            }
            return result;
        }

        private bool AreCollinear(XYZ p1, XYZ p2, XYZ p3)
        {
            if (p1.DistanceTo(p2) < 1e-5 || p2.DistanceTo(p3) < 1e-5) return true;
            XYZ v1 = (p2 - p1).Normalize();
            XYZ v2 = (p3 - p2).Normalize();
            return v1.DotProduct(v2) > COLLINEAR_DOT;
        }

        // ═══════════════════════════════════════════════════════════════
        // HELPERS
        // ═══════════════════════════════════════════════════════════════
        private GraphicsStyle GetLineStyleByName(Document doc, string styleName)
        {
            if (string.IsNullOrEmpty(styleName)) return null;
            Category linesCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
            if (linesCat == null) return null;

            foreach (Category subCat in linesCat.SubCategories)
            {
                if (subCat.Name.Equals(styleName, StringComparison.OrdinalIgnoreCase))
                    return subCat.GetGraphicsStyle(GraphicsStyleType.Projection);
            }
            return null;
        }
    }
}