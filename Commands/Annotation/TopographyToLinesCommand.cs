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
                // 1. Get all Revit links
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
                    TaskDialog.Show("HMV - Topography to Lines",
                        "No se encontraron vínculos de Revit en el documento.");
                    return Result.Failed;
                }

                // 2. NEW: Get all available Line Styles
                List<string> styleNames = new List<string>();
                Category linesCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
                if (linesCat != null)
                {
                    foreach (Category subCat in linesCat.SubCategories)
                    {
                        styleNames.Add(subCat.Name);
                    }
                }
                styleNames.Sort(); // Alphabetical order for the user

                // Show UI and pass BOTH lists
                TopographyToLinesWindow win = new TopographyToLinesWindow(linkNames, styleNames);

                if (win.ShowDialog() != true)
                    return Result.Cancelled;

                int selectedIndex = win.SelectedLinkIndex;
                if (selectedIndex < 0 || selectedIndex >= links.Count)
                    return Result.Cancelled;

                RevitLinkInstance selectedLink = links[selectedIndex];

                // Read user inputs from Window
                double userOff1 = win.Offset1;
                double userOff2 = win.Offset2;
                double userOff3 = win.Offset3;
                string userStyleName = win.SelectedLineStyle;

                // Process the selected link and pass inputs AND the style name
                int created = 0, skipped = 0;
                ProcessTopography(doc, selectedLink, userOff1, userOff2, userOff3, userStyleName, out created, out skipped);

                TaskDialog.Show("HMV - Topography to Lines",
                    $"Proceso completado:\n\n" +
                    $"Líneas creadas: {created}\n" +
                    $"Líneas omitidas: {skipped}");

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
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
            double userOffset1_m,
            double userOffset2_m,
            double userOffset3_m,
            string userStyleName, // <-- New Parameter
            out int created,
            out int skipped)
        {
            created = 0;
            skipped = 0;

            View activeView = doc.ActiveView;

            // Section plane
            XYZ origin = activeView.Origin;
            XYZ viewDir = activeView.ViewDirection;

            // Get linked document
            Document linkDoc = link.GetLinkDocument();
            if (linkDoc == null)
            {
                TaskDialog.Show("Error",
                    "No se pudo acceder al documento vinculado.");
                return;
            }

            Transform transform = link.GetTotalTransform();

            // Collect topography meshes
            List<Mesh> meshes = new List<Mesh>();
            FilteredElementCollector topoCollector =
                new FilteredElementCollector(linkDoc)
                .OfCategory(BuiltInCategory.OST_Topography)
                .WhereElementIsNotElementType();

            Options opt = new Options();
            foreach (Element topo in topoCollector)
            {
                GeometryElement geom = topo.get_Geometry(opt);
                if (geom == null) continue;

                foreach (GeometryObject gObj in geom)
                {
                    if (gObj is Mesh mesh)
                        meshes.Add(mesh);
                }
            }

            // Extract intersection segments
            List<Tuple<XYZ, XYZ>> rawSegments =
                ExtractIntersectionSegments(meshes, transform, origin, viewDir);

            // Merge micro-lines into fewer, longer lines
            List<Tuple<XYZ, XYZ>> segments = MergeCollinearSegments(rawSegments);

            // --- MULTIPLE OFFSET SETUP ---
            double initialOffset_mm = 100.0; // 100 millimeters
            double mmToFeet = 1.0 / 304.8;

            // Calculate absolute distances from the base line
            double dist1 = initialOffset_mm * mmToFeet;
            double dist2 = dist1 + (userOffset1_m * mmToFeet);
            double dist3 = dist2 + (userOffset2_m * mmToFeet);
            double dist4 = dist1 + (userOffset3_m * mmToFeet);

            // Create Transforms
            XYZ dir = activeView.UpDirection;
            Transform t1 = Transform.CreateTranslation(dir.Multiply(dist1));
            Transform t2 = Transform.CreateTranslation(dir.Multiply(dist2));
            Transform t3 = Transform.CreateTranslation(dir.Multiply(dist3));
            Transform t4 = Transform.CreateTranslation(dir.Multiply(dist4));

            // Create detail lines
            using (Transaction tx = new Transaction(doc, "HMV - Topography Base and Offsets"))
            {
                tx.Start();

                // Get styles before the loop
                GraphicsStyle thinStyle = GetLineStyleByName(doc, "<Thin Lines>");
                if (thinStyle == null) thinStyle = GetLineStyleByName(doc, "<Líneas finas>"); // Fallback for Spanish Revit

                GraphicsStyle userSelectedStyle = GetLineStyleByName(doc, userStyleName);

                foreach (Tuple<XYZ, XYZ> seg in segments)
                {
                    try
                    {
                        // 0. Base Topography Line (Uses <Thin Lines>)
                        Line baseLine = Line.CreateBound(seg.Item1, seg.Item2);
                        DetailCurve baseCurve = doc.Create.NewDetailCurve(activeView, baseLine);
                        if (thinStyle != null) baseCurve.LineStyle = thinStyle;
                        created++;

                        // 1. First Offset (Fixed 10 cm - Uses <Thin Lines>)
                        Line line1 = baseLine.CreateTransformed(t1) as Line;
                        DetailCurve dc1 = doc.Create.NewDetailCurve(activeView, line1);
                        if (thinStyle != null) dc1.LineStyle = thinStyle;
                        created++;

                        // 2. Second Offset (10cm + User Val 1 - Uses User Selection)
                        Line line2 = baseLine.CreateTransformed(t2) as Line;
                        DetailCurve dc2 = doc.Create.NewDetailCurve(activeView, line2);
                        if (userSelectedStyle != null) dc2.LineStyle = userSelectedStyle;
                        created++;

                        // 3. Third Offset (10cm + User Val 2 - Uses User Selection)
                        Line line3 = baseLine.CreateTransformed(t3) as Line;
                        DetailCurve dc3 = doc.Create.NewDetailCurve(activeView, line3);
                        if (userSelectedStyle != null) dc3.LineStyle = userSelectedStyle;
                        created++;

                        // 4. Fourth Offset (10cm + User Val 3 - Uses User Selection)
                        Line line4 = baseLine.CreateTransformed(t4) as Line;
                        DetailCurve dc4 = doc.Create.NewDetailCurve(activeView, line4);
                        if (userSelectedStyle != null) dc4.LineStyle = userSelectedStyle;
                        created++;
                    }
                    catch
                    {
                        skipped++;
                    }
                }

                tx.Commit();
            }
        }

        private GraphicsStyle GetLineStyleByName(Document doc, string styleName)
        {
            Category linesCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
            if (linesCat != null)
            {
                foreach (Category subCat in linesCat.SubCategories)
                {
                    if (subCat.Name.Equals(styleName, StringComparison.OrdinalIgnoreCase))
                    {
                        return subCat.GetGraphicsStyle(GraphicsStyleType.Projection);
                    }
                }
            }
            return null;
        }

        private List<Tuple<XYZ, XYZ>> ExtractIntersectionSegments(
            List<Mesh> meshes,
            Transform transform,
            XYZ origin,
            XYZ normal)
        {
            List<Tuple<XYZ, XYZ>> segments = new List<Tuple<XYZ, XYZ>>();

            foreach (Mesh mesh in meshes)
            {
                for (int i = 0; i < mesh.NumTriangles; i++)
                {
                    MeshTriangle tri = mesh.get_Triangle(i);

                    // Transform vertices
                    XYZ[] verts = new XYZ[3];
                    for (int j = 0; j < 3; j++)
                        verts[j] = transform.OfPoint(tri.get_Vertex(j));

                    // Compute signed distances
                    double[] dists = new double[3];
                    for (int j = 0; j < 3; j++)
                        dists[j] = SignedDistance(verts[j], origin, normal);

                    // Find intersection points
                    List<XYZ> intPts = new List<XYZ>();
                    int[,] edges = { { 0, 1 }, { 1, 2 }, { 2, 0 } };

                    for (int e = 0; e < 3; e++)
                    {
                        int e0 = edges[e, 0];
                        int e1 = edges[e, 1];
                        double d0 = dists[e0];
                        double d1 = dists[e1];

                        if (d0 * d1 < 0)
                        {
                            // Edge crosses plane
                            intPts.Add(Interpolate(verts[e0], verts[e1], d0, d1));
                        }
                        else if (Math.Abs(d0) < 1e-9)
                        {
                            // Vertex on plane
                            intPts.Add(verts[e0]);
                        }
                    }

                    // Remove duplicates
                    List<XYZ> unique = new List<XYZ>();
                    foreach (XYZ p in intPts)
                    {
                        bool isDup = false;
                        foreach (XYZ u in unique)
                        {
                            if (p.DistanceTo(u) < 1e-6)
                            {
                                isDup = true;
                                break;
                            }
                        }
                        if (!isDup) unique.Add(p);
                    }

                    // Add valid segment
                    if (unique.Count == 2 &&
                        unique[0].DistanceTo(unique[1]) > 1e-6)
                    {
                        segments.Add(Tuple.Create(unique[0], unique[1]));
                    }
                }
            }

            return segments;
        }

        private double SignedDistance(XYZ point, XYZ origin, XYZ normal)
        {
            XYZ vec = point - origin;
            return vec.DotProduct(normal);
        }

        private XYZ Interpolate(XYZ v0, XYZ v1, double d0, double d1)
        {
            double t = d0 / (d0 - d1);
            return new XYZ(
                v0.X + t * (v1.X - v0.X),
                v0.Y + t * (v1.Y - v0.Y),
                v0.Z + t * (v1.Z - v0.Z));
        }

        private List<Tuple<XYZ, XYZ>> MergeCollinearSegments(List<Tuple<XYZ, XYZ>> inputSegments)
        {
            if (inputSegments == null || inputSegments.Count < 2)
                return inputSegments;

            List<Tuple<XYZ, XYZ>> result = new List<Tuple<XYZ, XYZ>>();
            List<Tuple<XYZ, XYZ>> pool = new List<Tuple<XYZ, XYZ>>(inputSegments);

            double tol = 1e-4; // Tolerance for endpoints touching

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

                        // Case 1: Candidate connects to our line's END point
                        if (candidate.Item1.DistanceTo(end) < tol)
                        {
                            if (AreCollinear(start, end, candidate.Item2))
                            {
                                end = candidate.Item2; // Extend our line
                                pool.RemoveAt(i);
                                extended = true; break;
                            }
                        }
                        else if (candidate.Item2.DistanceTo(end) < tol)
                        {
                            if (AreCollinear(start, end, candidate.Item1))
                            {
                                end = candidate.Item1; // Extend our line
                                pool.RemoveAt(i);
                                extended = true; break;
                            }
                        }
                        // Case 2: Candidate connects to our line's START point
                        else if (candidate.Item2.DistanceTo(start) < tol)
                        {
                            if (AreCollinear(candidate.Item1, start, end))
                            {
                                start = candidate.Item1; // Extend backwards
                                pool.RemoveAt(i);
                                extended = true; break;
                            }
                        }
                        else if (candidate.Item1.DistanceTo(start) < tol)
                        {
                            if (AreCollinear(candidate.Item2, start, end))
                            {
                                start = candidate.Item2; // Extend backwards
                                pool.RemoveAt(i);
                                extended = true; break;
                            }
                        }
                    }
                }

                // Only add to result if the merged line is long enough for Revit (avoiding errors)
                if (start.DistanceTo(end) > 0.005)
                {
                    result.Add(new Tuple<XYZ, XYZ>(start, end));
                }
            }

            return result;
        }

        private bool AreCollinear(XYZ p1, XYZ p2, XYZ p3)
        {
            if (p1.DistanceTo(p2) < 1e-5 || p2.DistanceTo(p3) < 1e-5) return true;

            XYZ v1 = (p2 - p1).Normalize();
            XYZ v2 = (p3 - p2).Normalize();

            // Checks if the lines are pointing in the same direction.
            // 0.999 gives a ~2.5 degree angle tolerance so gently sloping terrain
            // also gets smoothed into longer lines.
            return v1.DotProduct(v2) > 0.99999;
        }
    }
}