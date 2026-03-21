using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

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
                // ── Collect candidate ImportInstances (imported + linked) ──
                List<ImportInstance> allImports = new FilteredElementCollector(doc)
                    .OfClass(typeof(ImportInstance))
                    .Cast<ImportInstance>()
                    .ToList();

                if (allImports.Count == 0)
                {
                    TaskDialog.Show("HMV - 3D DWG to Shape",
                        "No imported or linked DWG instances found in the project.");
                    return Result.Cancelled;
                }

                // ── Config window ──
                var configItems = allImports.Select(ii => new Dwg3DImportItem
                {
                    Id = ii.Id,
                    Name = (ii.Category?.Name ?? "DWG")
                         + (ii.IsLinked ? "  (Link)" : "  (Import)")
                         + "  [id " + ii.Id.IntegerValue + "]",
                    IsLinked = ii.IsLinked
                }).ToList();

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
                double meshBBoxMm = win.MeshBBoxMm;
                int decimateFactor = win.DecimateFactor;
                int selectedBic = win.SelectedCategoryBic;
                ElementId importId = win.SelectedImportId;
                bool deleteOriginal = win.DeleteOriginal;
                string dsName = win.ShapeName;

                double thresholdFt3 = (thresholdCm3 / 1000.0) / 0.0283168;
                double meshBBoxFt = meshBBoxMm / 304.8;  // mm → ft

                ImportInstance imp = doc.GetElement(importId) as ImportInstance;
                if (imp == null) { message = "Selected import not found."; return Result.Failed; }

                // ══════════════════════════════════════════
                //  PROGRESS WINDOW (modeless)
                // ══════════════════════════════════════════
                Dwg3DProgressWindow prog = new Dwg3DProgressWindow();
                prog.Show();

                try
                {
                    // ── Phase 1: Extract geometry ──
                    prog.UpdatePhase("Phase 1/3 — Extracting geometry...");
                    prog.UpdateDetail("Reading DWG geometry tree...");
                    prog.SetIndeterminate();

                    Options opt = new Options
                    {
                        ComputeReferences = true,
                        DetailLevel = ViewDetailLevel.Fine
                    };

                    GeometryElement geoElem = imp.get_Geometry(opt);
                    if (geoElem == null)
                    {
                        prog.Close();
                        message = "No geometry found.";
                        return Result.Failed;
                    }

                    List<Solid> solids = new List<Solid>();
                    List<Mesh> meshes = new List<Mesh>();
                    int skippedSmall = 0;
                    int geoCount = 0;

                    ExtractGeometry(geoElem, solids, meshes, thresholdFt3,
                                    ref skippedSmall, ref geoCount, prog);

                    if (prog.IsCancelled) { prog.Close(); return Result.Cancelled; }

                    prog.UpdateDetail(
                        $"Done: {solids.Count} solids, {meshes.Count} meshes, " +
                        $"{skippedSmall} skipped (below {thresholdCm3} cm³)");

                    // ── Phase 1b: Filter meshes by bounding box ──
                    int meshSkippedBBox = 0;
                    if (meshBBoxFt > 0 && meshes.Count > 0)
                    {
                        prog.UpdatePhase("Phase 1b — Filtering small meshes...");
                        List<Mesh> kept = new List<Mesh>();
                        for (int i = 0; i < meshes.Count; i++)
                        {
                            double diag = MeshBBoxDiagonal(meshes[i]);
                            if (diag >= meshBBoxFt)
                                kept.Add(meshes[i]);
                            else
                                meshSkippedBBox++;
                        }
                        prog.UpdateDetail(
                            $"Mesh filter: {kept.Count} kept, {meshSkippedBBox} skipped " +
                            $"(bbox < {meshBBoxMm} mm)");
                        meshes = kept;
                    }

                    // ── Phase 2: Convert meshes ──
                    prog.UpdatePhase($"Phase 2/3 — Converting meshes (decimate ×{decimateFactor})...");

                    List<GeometryObject> geoList = new List<GeometryObject>();
                    foreach (Solid s in solids) geoList.Add(s);

                    int meshConverted = 0;
                    int meshFailed = 0;
                    int totalTriangles = 0;
                    int degenerateSkipped = 0;
                    List<string> meshErrors = new List<string>();

                    if (meshes.Count > 0)
                    {
                        prog.SetDeterminate(meshes.Count);

                        for (int mi = 0; mi < meshes.Count; mi++)
                        {
                            if (prog.IsCancelled) break;

                            Mesh m = meshes[mi];
                            totalTriangles += m.NumTriangles;

                            prog.UpdateProgress(mi + 1, meshes.Count);
                            prog.UpdateDetail(
                                $"Mesh {mi + 1}/{meshes.Count}  " +
                                $"({m.NumTriangles} triangles)  " +
                                $"— OK: {meshConverted}, Failed: {meshFailed}");

                            try
                            {
                                var result = ConvertMesh(m, decimateFactor, out int degenCount);
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
                                    meshErrors.Add($"Mesh {mi + 1} ({m.NumTriangles} tri): builder returned 0 objects");
                                }
                            }
                            catch (Exception ex)
                            {
                                meshFailed++;
                                meshErrors.Add($"Mesh {mi + 1} ({m.NumTriangles} tri): {ex.Message}");
                            }
                        }
                    }

                    if (prog.IsCancelled) { prog.Close(); return Result.Cancelled; }

                    if (geoList.Count == 0)
                    {
                        prog.Close();
                        TaskDialog.Show("HMV - 3D DWG to Shape",
                            "No valid geometry extracted.\n\n" +
                            $"Solids found: {solids.Count}\n" +
                            $"Meshes found: {meshes.Count}\n" +
                            $"Skipped (below {thresholdCm3} cm³): {skippedSmall}\n" +
                            $"Mesh errors: {meshFailed}");
                        return Result.Cancelled;
                    }

                    // ── Phase 3: Create DirectShape ──
                    prog.UpdatePhase("Phase 3/3 — Creating DirectShape...");
                    prog.UpdateDetail($"Building shape with {geoList.Count} geometry objects...");
                    prog.SetIndeterminate();

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

                        prog.Close();

                        // ── Final report ──
                        string report =
                            $"DirectShape created  (id {ds.Id.IntegerValue})\n\n" +
                            $"Solids kept:          {solids.Count}\n" +
                            $"Skipped (< {thresholdCm3} cm³):  {skippedSmall}\n" +
                            $"Meshes found:         {meshes.Count + meshSkippedBBox}\n" +
                            $"  → bbox skipped:     {meshSkippedBBox}  (< {meshBBoxMm} mm)\n" +
                            $"  → converted OK:     {meshConverted}\n" +
                            $"  → failed:           {meshFailed}\n" +
                            $"Total triangles:      {totalTriangles}\n" +
                            $"Decimate factor:      ×{decimateFactor}\n" +
                            $"Degenerate tri skip:  {degenerateSkipped}\n" +
                            $"Geometry objects:     {geoList.Count}\n" +
                            $"Original deleted:     {(deleteOriginal ? "Yes" : "No")}";

                        if (meshErrors.Count > 0)
                            report += "\n\nMesh errors:\n" +
                                string.Join("\n", meshErrors.Take(15).Select(e => "  • " + e));

                        TaskDialog td = new TaskDialog("HMV - 3D DWG to Shape");
                        td.MainInstruction = "DirectShape created successfully.";
                        td.MainContent = report;
                        td.Show();
                    }
                }
                catch
                {
                    prog.Close();
                    throw;
                }

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            { return Result.Cancelled; }
            catch (Exception ex)
            { message = ex.Message; return Result.Failed; }
        }

        // ══════════════════════════════════════════════════════
        //  Recursive geometry extraction (with progress)
        // ══════════════════════════════════════════════════════
        void ExtractGeometry(GeometryElement geoElem,
            List<Solid> solids, List<Mesh> meshes,
            double thresholdFt3, ref int skipped,
            ref int geoCount, Dwg3DProgressWindow prog)
        {
            foreach (GeometryObject g in geoElem)
            {
                if (prog.IsCancelled) return;

                geoCount++;
                if (geoCount % 50 == 0)
                    prog.UpdateDetail(
                        $"Scanning... {geoCount} objects  |  " +
                        $"{solids.Count} solids, {meshes.Count} meshes, {skipped} skipped");

                if (g is Solid solid)
                {
                    if (solid.Faces.Size == 0) continue;
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
                    GeometryElement instGeo = gi.GetInstanceGeometry();
                    if (instGeo != null)
                        ExtractGeometry(instGeo, solids, meshes, thresholdFt3,
                                        ref skipped, ref geoCount, prog);
                }
            }
        }

        // ══════════════════════════════════════════════════════
        //  Mesh → TessellatedShapeBuilder
        //
        //  Key fixes vs Dynamo/Python version:
        //  1. Target = AnyGeometry, Fallback = Mesh
        //  2. Proper List<XYZ> for TessellatedFace
        //  3. Degenerate triangle filtering
        //  4. Decimation: only every Nth triangle (1 = all)
        // ══════════════════════════════════════════════════════
        IList<GeometryObject> ConvertMesh(Mesh mesh, int decimateFactor, out int degenerateCount)
        {
            degenerateCount = 0;

            TessellatedShapeBuilder builder = new TessellatedShapeBuilder();
            builder.Target = TessellatedShapeBuilderTarget.AnyGeometry;
            builder.Fallback = TessellatedShapeBuilderFallback.Mesh;

            builder.OpenConnectedFaceSet(false);

            int validFaces = 0;
            int step = Math.Max(1, decimateFactor);

            for (int i = 0; i < mesh.NumTriangles; i += step)
            {
                MeshTriangle tri = mesh.get_Triangle(i);
                XYZ v0 = tri.get_Vertex(0);
                XYZ v1 = tri.get_Vertex(1);
                XYZ v2 = tri.get_Vertex(2);

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

        // ══════════════════════════════════════════════════════
        //  Mesh bounding box diagonal (in ft, same as Revit internal)
        //  Iterates vertices to find min/max XYZ, returns diagonal length
        // ══════════════════════════════════════════════════════
        double MeshBBoxDiagonal(Mesh mesh)
        {
            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;

            IList<XYZ> verts = mesh.Vertices;
            for (int i = 0; i < verts.Count; i++)
            {
                XYZ v = verts[i];
                if (v.X < minX) minX = v.X; if (v.X > maxX) maxX = v.X;
                if (v.Y < minY) minY = v.Y; if (v.Y > maxY) maxY = v.Y;
                if (v.Z < minZ) minZ = v.Z; if (v.Z > maxZ) maxZ = v.Z;
            }

            double dx = maxX - minX, dy = maxY - minY, dz = maxZ - minZ;
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }

        bool IsDegenerate(XYZ a, XYZ b, XYZ c)
        {
            const double TOL = 1e-9;
            XYZ cross = (b - a).CrossProduct(c - a);
            return cross.GetLength() < TOL;
        }
    }

    // ── Plain data classes ──
    public class Dwg3DImportItem
    {
        public ElementId Id { get; set; }
        public string Name { get; set; }
        public bool IsLinked { get; set; }
        public override string ToString() => Name;
    }

    public class Dwg3DCategoryChoice
    {
        public string Label { get; set; }
        public int BicInt { get; set; }
        public override string ToString() => Label;
    }
}