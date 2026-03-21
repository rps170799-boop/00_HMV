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
    public class Dwg3DToShapeCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            try
            {
                // ── Collect candidate ImportInstances (3D only) ──
                List<ImportInstance> allImports = new FilteredElementCollector(doc)
                    .OfClass(typeof(ImportInstance))
                    .Cast<ImportInstance>()
                    .Where(ii => !ii.IsLinked)
                    .ToList();

                if (allImports.Count == 0)
                {
                    TaskDialog.Show("HMV - 3D DWG to Shape",
                        "No imported DWG instances found in the project.\n" +
                        "Import (not link) a 3D DWG first.");
                    return Result.Cancelled;
                }

                // ── Config window ──
                var configItems = allImports.Select(ii => new Dwg3DImportItem
                {
                    Id = ii.Id,
                    Name = (ii.Category?.Name ?? "Import") + "  [id " + ii.Id.IntegerValue + "]"
                }).ToList();

                // Category choices (plain data, no Revit types in window)
                var catChoices = new List<Dwg3DCategoryChoice>
                {
                    new Dwg3DCategoryChoice { Label = "Mechanical Equipment", BicInt = (int)BuiltInCategory.OST_MechanicalEquipment },
                    new Dwg3DCategoryChoice { Label = "Generic Models",       BicInt = (int)BuiltInCategory.OST_GenericModel },
                    new Dwg3DCategoryChoice { Label = "Structural Framing",   BicInt = (int)BuiltInCategory.OST_StructuralFraming },
                    new Dwg3DCategoryChoice { Label = "Structural Columns",   BicInt = (int)BuiltInCategory.OST_StructuralColumns },
                    new Dwg3DCategoryChoice { Label = "Electrical Equipment",  BicInt = (int)BuiltInCategory.OST_ElectricalEquipment },
                    new Dwg3DCategoryChoice { Label = "Specialty Equipment",   BicInt = (int)BuiltInCategory.OST_SpecialityEquipment },
                };

                Dwg3DToShapeWindow win = new Dwg3DToShapeWindow(configItems, catChoices);
                if (win.ShowDialog() != true) return Result.Cancelled;

                double thresholdCm3 = win.ThresholdCm3;
                int selectedBic = win.SelectedCategoryBic;
                ElementId importId = win.SelectedImportId;
                bool deleteOriginal = win.DeleteOriginal;
                string dsName = win.ShapeName;

                // threshold: cm³ → ft³
                double thresholdFt3 = (thresholdCm3 / 1000.0) / 0.0283168;

                ImportInstance imp = doc.GetElement(importId) as ImportInstance;
                if (imp == null) { message = "Selected import not found."; return Result.Failed; }

                // ── Extract geometry ──
                Options opt = new Options
                {
                    ComputeReferences = true,
                    DetailLevel = ViewDetailLevel.Fine
                };

                GeometryElement geoElem = imp.get_Geometry(opt);
                if (geoElem == null) { message = "No geometry found."; return Result.Failed; }

                List<Solid> solids = new List<Solid>();
                List<Mesh> meshes = new List<Mesh>();
                int skippedSmall = 0;

                ExtractGeometry(geoElem, solids, meshes, thresholdFt3, ref skippedSmall);

                // ── Convert meshes → GeometryObjects via TessellatedShapeBuilder ──
                List<GeometryObject> geoList = new List<GeometryObject>();
                foreach (Solid s in solids) geoList.Add(s);

                int meshConverted = 0;
                int meshFailed = 0;
                int totalTriangles = 0;
                int degenerateSkipped = 0;
                List<string> meshErrors = new List<string>();

                foreach (Mesh m in meshes)
                {
                    try
                    {
                        totalTriangles += m.NumTriangles;
                        var result = ConvertMesh(m, out int degenCount);
                        degenerateSkipped += degenCount;

                        if (result != null && result.Count > 0)
                        {
                            foreach (GeometryObject go in result)
                                geoList.Add(go);
                            meshConverted++;
                        }
                        else
                        {
                            meshFailed++;
                            meshErrors.Add($"Mesh ({m.NumTriangles} tri): builder returned 0 objects");
                        }
                    }
                    catch (Exception ex)
                    {
                        meshFailed++;
                        meshErrors.Add($"Mesh ({m.NumTriangles} tri): {ex.Message}");
                    }
                }

                if (geoList.Count == 0)
                {
                    TaskDialog.Show("HMV - 3D DWG to Shape",
                        "No valid geometry extracted.\n\n" +
                        $"Solids found: {solids.Count}\n" +
                        $"Meshes found: {meshes.Count}\n" +
                        $"Skipped (below {thresholdCm3} cm³): {skippedSmall}\n" +
                        $"Mesh errors: {meshFailed}");
                    return Result.Cancelled;
                }

                // ── Create DirectShape ──
                ElementId catId = new ElementId((BuiltInCategory)selectedBic);

                using (Transaction tx = new Transaction(doc, "HMV - 3D DWG to DirectShape"))
                {
                    tx.Start();

                    DirectShape ds = DirectShape.CreateElement(doc, catId);
                    ds.SetShape(geoList);
                    if (!string.IsNullOrWhiteSpace(dsName))
                        ds.SetName(dsName);

                    if (deleteOriginal)
                        doc.Delete(importId);

                    tx.Commit();

                    // ── Report ──
                    string report =
                        $"DirectShape created  (id {ds.Id.IntegerValue})\n\n" +
                        $"Solids kept:          {solids.Count}\n" +
                        $"Skipped (< {thresholdCm3} cm³):  {skippedSmall}\n" +
                        $"Meshes found:         {meshes.Count}\n" +
                        $"  → converted OK:     {meshConverted}\n" +
                        $"  → failed:           {meshFailed}\n" +
                        $"Total triangles:      {totalTriangles}\n" +
                        $"Degenerate tri skip:  {degenerateSkipped}\n" +
                        $"Geometry objects:     {geoList.Count}\n" +
                        $"Original deleted:     {(deleteOriginal ? "Yes" : "No")}";

                    if (meshErrors.Count > 0)
                        report += "\n\nMesh errors:\n" + string.Join("\n", meshErrors.Take(10).Select(e => "  • " + e));

                    TaskDialog td = new TaskDialog("HMV - 3D DWG to Shape");
                    td.MainInstruction = "DirectShape created successfully.";
                    td.MainContent = report;
                    td.Show();
                }

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            { return Result.Cancelled; }
            catch (Exception ex)
            { message = ex.Message; return Result.Failed; }
        }

        // ══════════════════════════════════════════════════════
        //  Recursive geometry extraction
        // ══════════════════════════════════════════════════════
        void ExtractGeometry(GeometryElement geoElem,
            List<Solid> solids, List<Mesh> meshes,
            double thresholdFt3, ref int skipped)
        {
            foreach (GeometryObject g in geoElem)
            {
                if (g is Solid solid)
                {
                    if (solid.Faces.Size == 0) continue;  // skip invalid
                    try
                    {
                        if (solid.Volume > thresholdFt3)
                            solids.Add(solid);
                        else
                            skipped++;
                    }
                    catch { skipped++; }
                }
                else if (g is Mesh mesh)
                {
                    if (mesh.NumTriangles > 0)
                        meshes.Add(mesh);
                }
                else if (g is GeometryInstance gi)
                {
                    // GetInstanceGeometry() returns geometry in project coords
                    GeometryElement instGeo = gi.GetInstanceGeometry();
                    if (instGeo != null)
                        ExtractGeometry(instGeo, solids, meshes, thresholdFt3, ref skipped);
                }
            }
        }

        // ══════════════════════════════════════════════════════
        //  Mesh → TessellatedShapeBuilder (THE FIX)
        //
        //  Key differences vs the Dynamo/Python version:
        //  1. Target = AnyGeometry, Fallback = Mesh
        //     (Python had no Target/Fallback → defaulted to Solid
        //      which silently fails on open/non-watertight meshes)
        //  2. Proper List<XYZ> instead of Python list
        //     (CPython3 list doesn't auto-convert to IList<XYZ>)
        //  3. Degenerate triangle filtering (zero-area faces crash builder)
        // ══════════════════════════════════════════════════════
        IList<GeometryObject> ConvertMesh(Mesh mesh, out int degenerateCount)
        {
            degenerateCount = 0;

            TessellatedShapeBuilder builder = new TessellatedShapeBuilder();

            // ★ THIS is what was missing in the Python version ★
            builder.Target = TessellatedShapeBuilderTarget.AnyGeometry;
            builder.Fallback = TessellatedShapeBuilderFallback.Mesh;

            builder.OpenConnectedFaceSet(false); // false = open shell (not solid)

            int validFaces = 0;
            for (int i = 0; i < mesh.NumTriangles; i++)
            {
                MeshTriangle tri = mesh.get_Triangle(i);
                XYZ v0 = tri.get_Vertex(0);
                XYZ v1 = tri.get_Vertex(1);
                XYZ v2 = tri.get_Vertex(2);

                // Skip degenerate triangles (collinear or duplicate vertices)
                if (IsDegenerate(v0, v1, v2))
                {
                    degenerateCount++;
                    continue;
                }

                List<XYZ> pts = new List<XYZ> { v0, v1, v2 };
                TessellatedFace face = new TessellatedFace(pts, ElementId.InvalidElementId);
                builder.AddFace(face);
                validFaces++;
            }

            builder.CloseConnectedFaceSet();

            if (validFaces == 0) return null;

            builder.Build();
            TessellatedShapeBuilderResult result = builder.GetBuildResult();
            IList<GeometryObject> geos = result.GetGeometricalObjects();

            return (geos != null && geos.Count > 0) ? geos : null;
        }

        bool IsDegenerate(XYZ a, XYZ b, XYZ c)
        {
            const double TOL = 1e-9; // ft²
            XYZ cross = (b - a).CrossProduct(c - a);
            return cross.GetLength() < TOL;
        }
    }

    // ── Plain data classes (no Revit references) ──
    public class Dwg3DImportItem
    {
        public ElementId Id { get; set; }
        public string Name { get; set; }
        public override string ToString() => Name;
    }

    public class Dwg3DCategoryChoice
    {
        public string Label { get; set; }
        public int BicInt { get; set; }
        public override string ToString() => Label;
    }
}