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
                // Get all Revit links
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

                // Show UI to select link
                TopographyToLinesWindow win = new TopographyToLinesWindow(linkNames);

                if (win.ShowDialog() != true)
                    return Result.Cancelled;

                int selectedIndex = win.SelectedLinkIndex;
                if (selectedIndex < 0 || selectedIndex >= links.Count)
                    return Result.Cancelled;

                RevitLinkInstance selectedLink = links[selectedIndex];

                // Process the selected link
                int created = 0, skipped = 0;
                ProcessTopography(doc, selectedLink, out created, out skipped);

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
            List<Tuple<XYZ, XYZ>> segments = 
                ExtractIntersectionSegments(meshes, transform, origin, viewDir);

            // Create detail lines
            using (Transaction tx = new Transaction(doc, 
                "HMV - Topography to Lines"))
            {
                tx.Start();

                foreach (Tuple<XYZ, XYZ> seg in segments)
                {
                    try
                    {
                        Line line = Line.CreateBound(seg.Item1, seg.Item2);
                        doc.Create.NewDetailCurve(activeView, line);
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
    }
}