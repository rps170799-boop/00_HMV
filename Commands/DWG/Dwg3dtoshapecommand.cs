using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using IOPath = System.IO.Path;
using System.Linq;

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
                // ── Collect existing DirectShapes (for Mode C: migrate to family) ──
                List<DirectShape> allDirectShapes = new FilteredElementCollector(doc)
                    .OfClass(typeof(DirectShape))
                    .Cast<DirectShape>()
                    .Where(ds => ds.Category != null)
                    .ToList();

                if (allImports.Count == 0 && allDirectShapes.Count == 0)
                {
                    TaskDialog.Show("HMV - 3D DWG to Shape",
                        "No DWG imports or DirectShapes found in the project.");
                    return Result.Cancelled;
                }

                // ── Config window ──
                var configItems = allImports.Select(ii =>
                {
                    // Try to get the DWG filename from the CADLinkType
                    string baseName = "DWG_DirectShape";
                    try
                    {
                        CADLinkType cadType = doc.GetElement(ii.GetTypeId()) as CADLinkType;
                        if (cadType != null)
                        {
                            ExternalFileReference extRef = cadType.GetExternalFileReference();
                            if (extRef != null)
                            {
                                ModelPath mp = extRef.GetAbsolutePath();
                                string full = ModelPathUtils.ConvertModelPathToUserVisiblePath(mp);
                                if (!string.IsNullOrEmpty(full))
                                    baseName = System.IO.Path.GetFileNameWithoutExtension(full);
                            }
                        }
                        // Fallback: try the type's Name parameter
                        if (baseName == "DWG_DirectShape" && cadType != null)
                        {
                            string typeName = cadType.Name;
                            if (!string.IsNullOrWhiteSpace(typeName))
                                baseName = System.IO.Path.GetFileNameWithoutExtension(typeName);
                        }
                    }
                    catch { /* keep default */ }

                    return new Dwg3DImportItem
                    {
                        Id = ii.Id,
                        Name = (ii.Category?.Name ?? "DWG")
                             + (ii.IsLinked ? "  (Link)" : "  (Import)")
                             + "  [id " + ii.Id.IntegerValue + "]"
                             + "  — " + baseName,
                        BaseFileName = baseName,
                        IsLinked = ii.IsLinked
                    };
                }).ToList();

                var dsItems = allDirectShapes.Select(ds => new Dwg3DShapeItem
                {
                    Id = ds.Id,
                    Name = (string.IsNullOrWhiteSpace(ds.Name) ? "DirectShape" : ds.Name)
                 + "  [" + (ds.Category?.Name ?? "?") + "]"
                 + "  [id " + ds.Id.IntegerValue + "]"
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

                Dwg3DToShapeWindow win = new Dwg3DToShapeWindow(configItems, catChoices, dsItems);
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

                // ── Resolve source element based on mode ──
                ImportInstance imp = null;
                DirectShape sourceDs = null;

                if (win.Mode == Dwg3DOutputMode.MigrateDsToFamily)
                {
                    sourceDs = doc.GetElement(win.SelectedSourceDsId) as DirectShape;
                    if (sourceDs == null) { message = "Selected DirectShape not found."; return Result.Failed; }
                }
                else
                {
                    imp = doc.GetElement(importId) as ImportInstance;
                    if (imp == null) { message = "Selected import not found."; return Result.Failed; }
                }

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
                    GeometryElement geoElem = (win.Mode == Dwg3DOutputMode.MigrateDsToFamily)
                    ? sourceDs.get_Geometry(opt)
                    : imp.get_Geometry(opt);
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

                    // ── Phase 3: Create output ──
                    if (win.Mode == Dwg3DOutputMode.DirectShapeInProject)
                    {
                        // ── Phase 3: Create DirectShape ──
                        prog.UpdatePhase("Phase 3/3 — Creating DirectShape...");
                    prog.SetDeterminate(geoList.Count);

                    ElementId catId = new ElementId((BuiltInCategory)selectedBic);
                    int geoAdded = 0;
                    int geoRejected = 0;

                    using (Transaction tx = new Transaction(doc, "HMV - 3D DWG to DirectShape"))
                    {
                        tx.Start();

                        DirectShape ds = DirectShape.CreateElement(doc, catId);
                        if (!string.IsNullOrWhiteSpace(dsName))
                            ds.SetName(dsName);

                        // Add geometry one by one — skip any that fail validation
                        for (int gi = 0; gi < geoList.Count; gi++)
                        {
                            if (prog.IsCancelled) break;

                            if (gi % 10 == 0)
                            {
                                prog.UpdateProgress(gi + 1, geoList.Count);
                                prog.UpdateDetail(
                                    $"Adding geometry {gi + 1}/{geoList.Count}  " +
                                    $"— OK: {geoAdded}, Rejected: {geoRejected}");
                            }

                            try
                            {
                                // Pre-validate: skip zero-volume solids
                                if (geoList[gi] is Solid sol && sol.Faces.Size == 0)
                                {
                                    geoRejected++;
                                    continue;
                                }

                                ds.AppendShape(new List<GeometryObject> { geoList[gi] });
                                geoAdded++;
                            }
                            catch
                            {
                                geoRejected++;
                            }
                        }

                        if (prog.IsCancelled)
                        {
                            tx.RollBack();
                            prog.Close();
                            return Result.Cancelled;
                        }

                        if (geoAdded == 0)
                        {
                            tx.RollBack();
                            prog.Close();
                            TaskDialog.Show("HMV - 3D DWG to Shape",
                                "All geometry objects were rejected by Revit.\n" +
                                $"Total attempted: {geoList.Count}\n" +
                                $"Rejected: {geoRejected}");
                            return Result.Cancelled;
                        }

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
                            $"Geometry added:       {geoAdded}\n" +
                            $"Geometry rejected:    {geoRejected}\n" +
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
                    else
                    {
                        // Mode B (family from DWG) or Mode C (migrate DS to family)
                        // — family path, implemented in Chunks 3 & 4.
                        XYZ insertionPoint = ComputeInsertionPoint(geoList, imp, sourceDs);
                        CreateFamilyWithDirectShape(
                            commandData.Application.Application,
                            doc, geoList, insertionPoint,
                            selectedBic, dsName,
                            win.Mode, win.CreateReferencePlanes, win.CreateCoarseCube,
                            importId, sourceDs?.Id, deleteOriginal,
                            prog);
                        prog.Close();
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

        /// <summary>
        /// Returns the world-coordinate point where the family instance
        /// should be placed, so the family origin lines up with the
        /// original DWG/DirectShape location.
        /// Strategy: use the DWG's LocationPoint if available; otherwise
        /// fall back to the bounding-box center of the extracted geoList.
        /// </summary>
        XYZ ComputeInsertionPoint(
            List<GeometryObject> geoList,
            ImportInstance imp,
            DirectShape sourceDs)
        {
            // Prefer the native location of the source element
            if (imp != null && imp.Location is LocationPoint lp && lp.Point != null)
                return lp.Point;

            if (sourceDs != null)
            {
                BoundingBoxXYZ bb = sourceDs.get_BoundingBox(null);
                if (bb != null)
                    return (bb.Min + bb.Max) * 0.5;
            }

            // Fallback: compute bbox center from geoList itself
            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;

            foreach (GeometryObject g in geoList)
            {
                if (g is Solid s && s.Faces.Size > 0)
                {
                    BoundingBoxXYZ sb = null;
                    try { sb = s.GetBoundingBox(); } catch { }
                    if (sb != null)
                    {
                        // Solid.GetBoundingBox() is in solid-local coords, but for
                        // DWG-extracted solids the transform is identity in practice.
                        if (sb.Min.X < minX) minX = sb.Min.X;
                        if (sb.Min.Y < minY) minY = sb.Min.Y;
                        if (sb.Min.Z < minZ) minZ = sb.Min.Z;
                        if (sb.Max.X > maxX) maxX = sb.Max.X;
                        if (sb.Max.Y > maxY) maxY = sb.Max.Y;
                        if (sb.Max.Z > maxZ) maxZ = sb.Max.Z;
                    }
                }
                else if (g is Mesh m)
                {
                    IList<XYZ> verts = m.Vertices;
                    for (int i = 0; i < verts.Count; i++)
                    {
                        XYZ v = verts[i];
                        if (v.X < minX) minX = v.X; if (v.X > maxX) maxX = v.X;
                        if (v.Y < minY) minY = v.Y; if (v.Y > maxY) maxY = v.Y;
                        if (v.Z < minZ) minZ = v.Z; if (v.Z > maxZ) maxZ = v.Z;
                    }
                }
            }

            if (minX == double.MaxValue) return XYZ.Zero;
            return new XYZ((minX + maxX) * 0.5, (minY + maxY) * 0.5, (minZ + maxZ) * 0.5);
        }

        /// <summary>
        /// Creates a new family containing a DirectShape built from geoList,
        /// saves it to %TEMP%, loads it into the project, and places one
        /// instance at insertionPoint. Category falls back to Generic Model
        /// if the requested BIC is not allowed by the family template.
        /// </summary>
        void CreateFamilyWithDirectShape(
            Autodesk.Revit.ApplicationServices.Application app,
            Document projectDoc,
            List<GeometryObject> geoList,
            XYZ insertionPoint,
            int bicInt,
            string shapeName,
            Dwg3DOutputMode mode,
            bool createRefPlanes,
            bool createCoarseCube,
            ElementId originalImportId,
            ElementId sourceDsId,
            bool deleteOriginal,
            Dwg3DProgressWindow prog)
        {
            // ── Phase 3a: Resolve template ──
            prog.UpdatePhase("Phase 3/6 — Resolving family template...");
            prog.SetIndeterminate();

            string templatePath = ResolveGenericModelTemplate(app);
            prog.UpdateDetail("Template resolved to: " + (templatePath ?? "<null>"));
            System.IO.Directory.GetFiles(app.FamilyTemplatePath, "Generic Model*.rft", System.IO.SearchOption.AllDirectories);
            if (templatePath == null)
            {
                prog.Close();
                TaskDialog.Show("HMV - Family Creation",
                    "Could not find a Generic Model family template (.rft).\n\n" +
                    "Searched in: " + app.FamilyTemplatePath);
                return;
            }

            // In-place replacement: the family's origin sits at the DWG's
            // LocationPoint, so placing the family instance back at that
            // same point puts geometry exactly where the DWG was.
            XYZ offset = -insertionPoint;
            Transform tx = Transform.CreateTranslation(offset);
           

            List<GeometryObject> localGeo = new List<GeometryObject>();
            int translateFailed = 0;
            foreach (GeometryObject g in geoList)
            {
                try
                {
                    if (g is Solid s)
                        localGeo.Add(SolidUtils.CreateTransformed(s, tx));
                    else if (g is Mesh m)
                        localGeo.Add(TranslateMesh(m, offset));
                    else
                        localGeo.Add(g);
                }
                catch { translateFailed++; }
            }

            // ── Phase 3c: Create family document ──
            prog.UpdatePhase("Phase 5/6 — Creating family document...");
            Document famDoc = null;
            string tempRfa = null;
            bool categoryFellBack = false;
            int famGeoAdded = 0, famGeoRejected = 0;

            try
            {
                famDoc = app.NewFamilyDocument(templatePath);
                if (famDoc == null)
                {
                    prog.Close();
                    TaskDialog.Show("HMV - Family Creation",
                        "NewFamilyDocument returned null. Template: " + templatePath);
                    return;
                }

                using (Transaction ftx = new Transaction(famDoc, "HMV - Build family"))
                {
                    ftx.Start();

                    // Try to set the requested family category, fall back to Generic Model
                    try
                    {
                        Category targetCat = Category.GetCategory(famDoc, (BuiltInCategory)bicInt);
                        if (targetCat != null)
                            famDoc.OwnerFamily.FamilyCategory = targetCat;
                        else
                            categoryFellBack = true;
                    }
                    catch { categoryFellBack = true; }

                    // DirectShape INSIDE the family — same API as project DS
                    ElementId dsCatId = new ElementId((BuiltInCategory)bicInt);
                    if (!DirectShape.IsValidCategoryId(dsCatId, famDoc))
                    {
                        dsCatId = new ElementId(BuiltInCategory.OST_GenericModel);
                        categoryFellBack = true;
                    }

                    DirectShape fds = DirectShape.CreateElement(famDoc, dsCatId);
                    if (!string.IsNullOrWhiteSpace(shapeName))
                        fds.SetName(shapeName);

                    for (int gi = 0; gi < localGeo.Count; gi++)
                    {
                        if (prog.IsCancelled) break;
                        try
                        {
                            if (localGeo[gi] is Solid sol && sol.Faces.Size == 0)
                            { famGeoRejected++; continue; }
                            fds.AppendShape(new List<GeometryObject> { localGeo[gi] });
                            famGeoAdded++;
                        }
                        catch { famGeoRejected++; }

                        if (gi % 10 == 0)
                            prog.UpdateDetail(
                                $"Family geometry {gi + 1}/{localGeo.Count}  " +
                                $"— OK: {famGeoAdded}, Rejected: {famGeoRejected}");
                    }

                    if (famGeoAdded == 0)
                    {
                        ftx.RollBack();
                        famDoc.Close(false);
                        prog.Close();
                        TaskDialog.Show("HMV - Family Creation",
                            "No geometry could be added to the family DirectShape.");
                        return;
                    }

                    // Chunk 4 — reference planes hook (no-op until Chunk 4 lands)
                    if (createRefPlanes)
                        CreateBoundingBoxReferencePlanes(famDoc, localGeo, prog);

                    if (createCoarseCube)
                        CreateCoarseVisibilityCube(famDoc, localGeo, prog);

                    ftx.Commit();
                }

                // ── Phase 3d: Save family to temp ──
                string tempDir = IOPath.Combine(IOPath.GetTempPath(), "HMVTools");
                Directory.CreateDirectory(tempDir);
                string safeName = shapeName;
                foreach (char c in IOPath.GetInvalidFileNameChars())
                    safeName = safeName.Replace(c, '_');
                tempRfa = IOPath.Combine(tempDir,
                    $"{safeName}_{DateTime.Now:yyyyMMdd_HHmmss}.rfa");

                SaveAsOptions sao = new SaveAsOptions { OverwriteExistingFile = true };
                famDoc.SaveAs(tempRfa, sao);
                // ── Pre-load cleanup: remove any existing instances of a family
                //     with the same name, so the re-run produces exactly 1 instance
                //     instead of stacking. LoadFamily itself will overwrite the
                //     family definition via TraceabilityLoadOptions.OnFamilyFound.
                int removedOldInstances = 0;
                {
                    Family existingFamily = new FilteredElementCollector(projectDoc)
                        .OfClass(typeof(Family))
                        .Cast<Family>()
                        .FirstOrDefault(f => f.Name == shapeName);

                    if (existingFamily != null)
                    {
                        // Collect instance IDs across all symbols of the family
                        HashSet<ElementId> symIds = new HashSet<ElementId>(
                            existingFamily.GetFamilySymbolIds());

                        List<ElementId> instanceIds = new FilteredElementCollector(projectDoc)
                            .OfClass(typeof(FamilyInstance))
                            .Cast<FamilyInstance>()
                            .Where(fi => fi.Symbol != null && symIds.Contains(fi.Symbol.Id))
                            .Select(fi => fi.Id)
                            .ToList();

                        if (instanceIds.Count > 0)
                        {
                            using (Transaction cleanupTx = new Transaction(
                                projectDoc, "HMV - Remove old family instances"))
                            {
                                cleanupTx.Start();
                                try
                                {
                                    ICollection<ElementId> deleted = projectDoc.Delete(instanceIds);
                                    removedOldInstances = deleted?.Count ?? 0;
                                    cleanupTx.Commit();
                                }
                                catch
                                {
                                    if (cleanupTx.HasStarted()) cleanupTx.RollBack();
                                    // Non-fatal: proceed with the load anyway
                                }
                            }
                        }
                    }
                }
                // ── Phase 3e: Load into project & place ──
                prog.UpdatePhase("Phase 6/6 — Loading family into project...");

                // CRITICAL: load from the in-memory famDoc, NOT from the .rfa path.
                // The path-based LoadFamily is unreliable in 2023 when SaveAs has
                // just touched the file. The in-memory pattern is what our
                // Textauditcommand uses successfully.
                //
                // Also: LoadFamily must be called OUTSIDE any open project
                // transaction — it manages its own internal transaction.

                Family loadedFamily = null;
                try
                {
                    loadedFamily = famDoc.LoadFamily(projectDoc, new TraceabilityLoadOptions());
                }
                catch (Exception loadEx)
                {
                    try { famDoc.Close(false); } catch { }
                    famDoc = null;
                    prog.Close();
                    TaskDialog.Show("HMV - Family Creation",
                        "LoadFamily threw an exception:\n\n" + loadEx.Message);
                    return;
                }

                // Now close the famDoc — it's been loaded into the project
                try { famDoc.Close(false); } catch { }
                famDoc = null;

                // Look up the loaded family by name. The filename (minus .rfa)
                // is the family's name in the project after loading.
                string expectedName = IOPath.GetFileNameWithoutExtension(tempRfa);
                loadedFamily = new FilteredElementCollector(projectDoc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .FirstOrDefault(f => f.Name == expectedName);

                // Second fallback: try the clean shapeName (for re-runs)
                if (loadedFamily == null)
                {
                    loadedFamily = new FilteredElementCollector(projectDoc)
                        .OfClass(typeof(Family))
                        .Cast<Family>()
                        .FirstOrDefault(f => f.Name == shapeName);
                }

                if (loadedFamily == null)
                {
                    prog.Close();
                    TaskDialog.Show("HMV - Family Creation",
                        "LoadFamily succeeded but the family could not be located " +
                        "by name in the project.\n\n" +
                        "Expected: " + expectedName + " or " + shapeName);
                    return;
                }

                // Rename family to user input; type name = user input + "01"
                using (Transaction ptx = new Transaction(projectDoc, "HMV - Place family instance"))
                {
                    ptx.Start();

                    try { loadedFamily.Name = shapeName; } catch { }

                    FamilySymbol symbol = null;
                    foreach (ElementId sid in loadedFamily.GetFamilySymbolIds())
                    {
                        symbol = projectDoc.GetElement(sid) as FamilySymbol;
                        if (symbol != null) break;
                    }
                    if (symbol == null)
                    {
                        ptx.RollBack();
                        prog.Close();
                        TaskDialog.Show("HMV - Family Creation",
                            "Family loaded but no FamilySymbol was found.");
                        return;
                    }

                    try { symbol.Name = shapeName + "01"; } catch { }
                    if (!symbol.IsActive) symbol.Activate();

                    FamilyInstance inst = projectDoc.Create.NewFamilyInstance(
                        insertionPoint, symbol,
                        Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                    

                    if (deleteOriginal)
                    {
                        if (mode == Dwg3DOutputMode.NewFamilyFromDwg && originalImportId != null)
                            projectDoc.Delete(originalImportId);
                        else if (mode == Dwg3DOutputMode.MigrateDsToFamily && sourceDsId != null)
                            projectDoc.Delete(sourceDsId);
                    }

                    ptx.Commit();
                }

                prog.Close();

                // ── Final report ──
                string report =
                    $"Family created and placed.\n\n" +
                    $"Family name:       {shapeName}\n" +
                    $"Type name:         {shapeName}01\n" +
                    $"Template used:     {templatePath}\n" +
                    $"Category fallback: {(categoryFellBack ? "Yes (Generic Model)" : "No")}\n" +
                    $"Geometry added:    {famGeoAdded}\n" +
                    $"Geometry rejected: {famGeoRejected}\n" +
                    $"Translate failed:  {translateFailed}\n" +
                    $"Old instances removed: {removedOldInstances}\n" +
                    $"Reference planes:  {(createRefPlanes ? "Created" : "Skipped")}\n" +
                    $"Coarse cube:       {(createCoarseCube ? "Created" : "Skipped")}\n" +
                    $"Insertion point:   ({insertionPoint.X:F2}, {insertionPoint.Y:F2}, {insertionPoint.Z:F2})\n" +
                    $"Temp .rfa:         {tempRfa}\n" +
                    $"Original deleted:  {(deleteOriginal ? "Yes" : "No")}";

                TaskDialog.Show("HMV - Family Creation", report);
            }
            catch (Exception ex)
            {
                try { famDoc?.Close(false); } catch { }
                prog.Close();
                TaskDialog.Show("HMV - Family Creation Error",
                    "Family creation failed:\n\n" + ex.Message +
                    "\n\nStack:\n" + ex.StackTrace);
            }
        }

        // ── Helper: resolve Generic Model template ──
        string ResolveGenericModelTemplate(Autodesk.Revit.ApplicationServices.Application app)
        {
            string tpDir = app.FamilyTemplatePath;
            if (string.IsNullOrEmpty(tpDir) || !Directory.Exists(tpDir))
                return null;

            // Try common filenames first (metric + imperial).
            // Note: we deliberately avoid "Adaptive" templates — adaptive families
            // don't support NewExtrusion with a simple XYZ sketch plane and
            // cause the Coarse cube creation to fail.
            string[] preferred = new[]
            {
                "Metric Generic Model.rft",
                "Generic Model.rft"
            };
            foreach (string name in preferred)
            {
                string full = IOPath.Combine(tpDir, name);
                if (File.Exists(full)) return full;
            }

            // Recursive fallback — but explicitly exclude Adaptive, Pattern,
            // Conceptual Mass, and Face-Based variants.
            try
            {
                string[] matches = Directory.GetFiles(
                    tpDir, "*Generic Model*.rft", SearchOption.AllDirectories);

                foreach (string m in matches)
                {
                    string fname = IOPath.GetFileName(m);
                    if (fname.IndexOf("Adaptive", StringComparison.OrdinalIgnoreCase) >= 0) continue;
                    if (fname.IndexOf("Pattern", StringComparison.OrdinalIgnoreCase) >= 0) continue;
                    if (fname.IndexOf("Face", StringComparison.OrdinalIgnoreCase) >= 0) continue;
                    if (fname.IndexOf("Mass", StringComparison.OrdinalIgnoreCase) >= 0) continue;
                    if (fname.IndexOf("based", StringComparison.OrdinalIgnoreCase) >= 0) continue;
                    if (fname.IndexOf("Casework", StringComparison.OrdinalIgnoreCase) >= 0) continue;
                    if (fname.IndexOf("Curtain", StringComparison.OrdinalIgnoreCase) >= 0) continue;
                    return m;
                }
            }
            catch { }

            return null;
        }

        // ── Helper: translate a Mesh by an offset (via TessellatedShapeBuilder) ──
        GeometryObject TranslateMesh(Mesh mesh, XYZ offset)
        {
            TessellatedShapeBuilder tb = new TessellatedShapeBuilder();
            tb.Target = TessellatedShapeBuilderTarget.AnyGeometry;
            tb.Fallback = TessellatedShapeBuilderFallback.Mesh;
            tb.OpenConnectedFaceSet(false);

            for (int i = 0; i < mesh.NumTriangles; i++)
            {
                MeshTriangle tri = mesh.get_Triangle(i);
                XYZ v0 = tri.get_Vertex(0) + offset;
                XYZ v1 = tri.get_Vertex(1) + offset;
                XYZ v2 = tri.get_Vertex(2) + offset;
                if (IsDegenerate(v0, v1, v2)) continue;
                tb.AddFace(new TessellatedFace(
                    new List<XYZ> { v0, v1, v2 }, ElementId.InvalidElementId));
            }

            tb.CloseConnectedFaceSet();
            tb.Build();
            var res = tb.GetBuildResult();
            var geos = res.GetGeometricalObjects();
            return (geos != null && geos.Count > 0) ? geos[0] : mesh;
        }

        // ── Chunk 4 stub: reference planes (implemented next) ──
        // ══════════════════════════════════════════════════════
        //  Chunk 4: Bounding-box reference planes
        //
        //  Creates 9 reference planes inside the family document,
        //  arranged as a "cage" around the geometry's bounding box:
        //    6 face planes  (Left, Right, Front, Back, Bottom, Top)
        //    3 center planes (CenterX, CenterY, CenterZ)
        //
        //  Each plane gets a semantic IsReference value so Revit's
        //  automatic dimensioning and alignment behaves correctly
        //  when the family is placed in a project.
        //
        //  MUST be called inside an open famDoc transaction.
        // ══════════════════════════════════════════════════════
        void CreateBoundingBoxReferencePlanes(
            Document famDoc,
            List<GeometryObject> localGeo,
            Dwg3DProgressWindow prog)
        {
            prog.UpdateDetail("Computing bounding box for reference planes...");

            // ── Compute local-coord bbox from geoList ──
            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
            bool any = false;

            foreach (GeometryObject g in localGeo)
            {
                if (g is Solid s && s.Faces.Size > 0)
                {
                    try
                    {
                        foreach (Edge edge in s.Edges)
                        {
                            foreach (XYZ v in edge.Tessellate())
                            {
                                if (v.X < minX) minX = v.X; if (v.X > maxX) maxX = v.X;
                                if (v.Y < minY) minY = v.Y; if (v.Y > maxY) maxY = v.Y;
                                if (v.Z < minZ) minZ = v.Z; if (v.Z > maxZ) maxZ = v.Z;
                                any = true;
                            }
                        }
                    }
                    catch { }
                }
                else if (g is Mesh m)
                {
                    IList<XYZ> verts = m.Vertices;
                    for (int i = 0; i < verts.Count; i++)
                    {
                        XYZ v = verts[i];
                        if (v.X < minX) minX = v.X; if (v.X > maxX) maxX = v.X;
                        if (v.Y < minY) minY = v.Y; if (v.Y > maxY) maxY = v.Y;
                        if (v.Z < minZ) minZ = v.Z; if (v.Z > maxZ) maxZ = v.Z;
                        any = true;
                    }
                }
            }

            if (!any)
            {
                prog.UpdateDetail("No bbox data — skipping reference planes.");
                return;
            }

            // ── Padding: 10% of diagonal, clamped to [500mm, 2000mm] ──
            double dx = maxX - minX, dy = maxY - minY, dz = maxZ - minZ;
            double diag = Math.Sqrt(dx * dx + dy * dy + dz * dz);
            double padMin = 500.0 / 304.8;   // 500 mm → ft
            double padMax = 2000.0 / 304.8;  // 2000 mm → ft
            double pad = Math.Max(padMin, Math.Min(padMax, diag * 0.1));

            // Midpoints
            double cx = (minX + maxX) * 0.5;
            double cy = (minY + maxY) * 0.5;
            double cz = (minZ + maxZ) * 0.5;

            // ── Locate reference views for plane creation ──
            // Family templates have standard views: Front, Back, Left, Right,
            // Ref. Level. We need one for the horizontal planes (plan view)
            // and one for the vertical planes (elevation view).
            View planView = null;
            View elevView = null;
            foreach (View v in new FilteredElementCollector(famDoc)
                .OfClass(typeof(View)).Cast<View>())
            {
                if (v.IsTemplate) continue;
                if (planView == null && v.ViewType == ViewType.FloorPlan) planView = v;
                if (elevView == null && v.ViewType == ViewType.Elevation) elevView = v;
                if (planView != null && elevView != null) break;
            }

            if (planView == null || elevView == null)
            {
                prog.UpdateDetail("Could not locate family plan/elevation views — reference planes skipped.");
                return;
            }

            int created = 0, failed = 0;

            // ── Helper local function: create one ref plane ──
            ReferencePlane MakePlane(XYZ bubbleEnd, XYZ freeEnd, XYZ cutVec, View view, string name)
            {
                try
                {
                    ReferencePlane rp = famDoc.FamilyCreate.NewReferencePlane(
                        bubbleEnd, freeEnd, cutVec, view);
                    if (rp != null)
                    {
                        try { rp.Name = name; } catch { }
                        created++;
                    }
                    return rp;
                }
                catch { failed++; return null; }
            }

            // ── Helper: set IsReference semantic value ──
            void SetIsRef(ReferencePlane rp, int isRefValue)
            {
                if (rp == null) return;
                try
                {
                    Parameter p = rp.get_Parameter(
                        BuiltInParameter.ELEM_REFERENCE_NAME);
                    // Fallback: IsReference is usually exposed as a built-in
                    // parameter on ReferencePlane — try the enum-backed one
                    Parameter p2 = rp.get_Parameter(
                        BuiltInParameter.ELEM_IS_REFERENCE);
                    if (p2 != null && !p2.IsReadOnly)
                        p2.Set(isRefValue);
                }
                catch { }
            }

            // IsReference enum values (Revit internal):
            //   0=Not a Reference, 1=Strong, 2=Weak,
            //   3=Left, 4=Center LR, 5=Right,
            //   6=Front, 7=Center FB, 8=Back,
            //   9=Bottom, 10=Center Elev, 11=Top
            const int IR_LEFT = 3, IR_CLR = 4, IR_RIGHT = 5;
            const int IR_FRONT = 6, IR_CFB = 7, IR_BACK = 8;
            const int IR_BOTTOM = 9, IR_CEL = 10, IR_TOP = 11;

            // ── Vertical planes (X-normal): Left, Right, CenterX ──
            // A vertical plane parallel to YZ: line along Y at some X, cutVec = Z
            // Created in the elevation view.
            XYZ padY = new XYZ(0, pad, 0);
            XYZ upZ = XYZ.BasisZ;

            ReferencePlane rpLeft = MakePlane(
                new XYZ(minX, minY - pad, cz),
                new XYZ(minX, maxY + pad, cz),
                upZ, elevView, "HMV_Left");
            SetIsRef(rpLeft, IR_LEFT);

            ReferencePlane rpRight = MakePlane(
                new XYZ(maxX, minY - pad, cz),
                new XYZ(maxX, maxY + pad, cz),
                upZ, elevView, "HMV_Right");
            SetIsRef(rpRight, IR_RIGHT);

            ReferencePlane rpCX = MakePlane(
                new XYZ(cx, minY - pad, cz),
                new XYZ(cx, maxY + pad, cz),
                upZ, elevView, "HMV_CenterX");
            SetIsRef(rpCX, IR_CLR);

            // ── Vertical planes (Y-normal): Front, Back, CenterY ──
            // A vertical plane parallel to XZ: line along X at some Y, cutVec = Z
            ReferencePlane rpFront = MakePlane(
                new XYZ(minX - pad, minY, cz),
                new XYZ(maxX + pad, minY, cz),
                upZ, elevView, "HMV_Front");
            SetIsRef(rpFront, IR_FRONT);

            ReferencePlane rpBack = MakePlane(
                new XYZ(minX - pad, maxY, cz),
                new XYZ(maxX + pad, maxY, cz),
                upZ, elevView, "HMV_Back");
            SetIsRef(rpBack, IR_BACK);

            ReferencePlane rpCY = MakePlane(
                new XYZ(minX - pad, cy, cz),
                new XYZ(maxX + pad, cy, cz),
                upZ, elevView, "HMV_CenterY");
            SetIsRef(rpCY, IR_CFB);

            // ── Horizontal planes (Z-normal): Bottom, Top, CenterZ ──
            // A horizontal plane parallel to XY: line along X at some Z, cutVec = Y
            // Created in the plan view.
            XYZ outY = XYZ.BasisY;

            ReferencePlane rpBottom = MakePlane(
                new XYZ(minX - pad, cy, minZ),
                new XYZ(maxX + pad, cy, minZ),
                outY, planView, "HMV_Bottom");
            SetIsRef(rpBottom, IR_BOTTOM);

            ReferencePlane rpTop = MakePlane(
                new XYZ(minX - pad, cy, maxZ),
                new XYZ(maxX + pad, cy, maxZ),
                outY, planView, "HMV_Top");
            SetIsRef(rpTop, IR_TOP);

            ReferencePlane rpCZ = MakePlane(
                new XYZ(minX - pad, cy, cz),
                new XYZ(maxX + pad, cy, cz),
                outY, planView, "HMV_CenterZ");
            SetIsRef(rpCZ, IR_CEL);

            prog.UpdateDetail(
                $"Reference planes: {created} created, {failed} failed  " +
                $"(pad {(pad * 304.8):F0} mm)");
        }
        // ══════════════════════════════════════════════════════
        //  Coarse-view annotation cube
        //
        //  Creates a solid extrusion that wraps the geometry's
        //  bounding box (+1 mm padding on each face), visible
        //  only at Coarse detail level. At Medium and Fine the
        //  cube disappears and the real geometry is revealed.
        //
        //  Purpose: gives BIM coordinators a clean box to
        //  dimension and tag against in Coarse views, while
        //  preserving full fidelity in detail views.
        //
        //  MUST be called inside an open famDoc transaction.
        // ══════════════════════════════════════════════════════
        void CreateCoarseVisibilityCube(
            Document famDoc,
            List<GeometryObject> localGeo,
            Dwg3DProgressWindow prog)
        {
            prog.UpdateDetail("Creating Coarse-view annotation cube...");

            // ── Compute bbox from geoList (same pattern as ref planes) ──
            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
            bool any = false;

            foreach (GeometryObject g in localGeo)
            {
                if (g is Solid s && s.Faces.Size > 0)
                {
                    try
                    {
                        foreach (Edge edge in s.Edges)
                        {
                            foreach (XYZ v in edge.Tessellate())
                            {
                                if (v.X < minX) minX = v.X; if (v.X > maxX) maxX = v.X;
                                if (v.Y < minY) minY = v.Y; if (v.Y > maxY) maxY = v.Y;
                                if (v.Z < minZ) minZ = v.Z; if (v.Z > maxZ) maxZ = v.Z;
                                any = true;
                            }
                        }
                    }
                    catch { }
                }
                else if (g is Mesh m)
                {
                    IList<XYZ> verts = m.Vertices;
                    for (int i = 0; i < verts.Count; i++)
                    {
                        XYZ v = verts[i];
                        if (v.X < minX) minX = v.X; if (v.X > maxX) maxX = v.X;
                        if (v.Y < minY) minY = v.Y; if (v.Y > maxY) maxY = v.Y;
                        if (v.Z < minZ) minZ = v.Z; if (v.Z > maxZ) maxZ = v.Z;
                        any = true;
                    }
                }
            }

            if (!any)
            {
                prog.UpdateDetail("No bbox data — skipping Coarse cube.");
                return;
            }

            // ── Pad by 5 mm on each face (avoids Z-fighting with real geometry) ──
            const double CUBE_PAD_FT = 5.0 / 304.8;  // 5 mm in feet
            minX -= CUBE_PAD_FT; minY -= CUBE_PAD_FT; minZ -= CUBE_PAD_FT;
            maxX += CUBE_PAD_FT; maxY += CUBE_PAD_FT; maxZ += CUBE_PAD_FT;

            try
            {
                // ── Build a rectangle profile on the XY plane at Z=minZ ──
                XYZ p0 = new XYZ(minX, minY, minZ);
                XYZ p1 = new XYZ(maxX, minY, minZ);
                XYZ p2 = new XYZ(maxX, maxY, minZ);
                XYZ p3 = new XYZ(minX, maxY, minZ);

                CurveArray rect = new CurveArray();
                rect.Append(Line.CreateBound(p0, p1));
                rect.Append(Line.CreateBound(p1, p2));
                rect.Append(Line.CreateBound(p2, p3));
                rect.Append(Line.CreateBound(p3, p0));

                CurveArrArray profile = new CurveArrArray();
                profile.Append(rect);

                // ── Sketch plane on the base (Z=minZ), normal = +Z ──
                Plane basePlane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, p0);
                SketchPlane sketchPlane = SketchPlane.Create(famDoc, basePlane);

                // ── Extrude from 0 to (maxZ-minZ) along the plane's +Z normal ──
                double height = maxZ - minZ;
                Extrusion cube = famDoc.FamilyCreate.NewExtrusion(
                    true,            // isSolid
                    profile,
                    sketchPlane,
                    height);

                if (cube == null)
                {
                    prog.UpdateDetail("Coarse cube: NewExtrusion returned null.");
                    return;
                }

                try { cube.Name = "HMV_CoarseCube"; } catch { }

                // ── Visibility: visible only at Coarse ──
                // FamilyElementVisibility constructor takes a FamilyElementVisibilityType:
                //   Model = visible in 3D + all 2D view orientations.
                // Then we toggle the detail-level flags.
                FamilyElementVisibility vis = new FamilyElementVisibility(
                    FamilyElementVisibilityType.Model);
                vis.IsShownInCoarse = true;
                vis.IsShownInMedium = false;
                vis.IsShownInFine = false;

                
                cube.SetVisibility(vis);

                prog.UpdateDetail(
                    $"Coarse cube: created  (pad 1 mm, height " +
                    $"{(height * 304.8):F0} mm)");
            }
            catch (Exception ex)
            {
                prog.UpdateDetail("Coarse cube failed: " + ex.Message);
            }
        }
    }



    // ── Plain data classes ──
    public class Dwg3DImportItem
    {
        public ElementId Id { get; set; }
        public string Name { get; set; }
        public string BaseFileName { get; set; }
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