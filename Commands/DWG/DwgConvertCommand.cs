using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using ACadSharp;
using ACadSharp.IO;
using AcadText = ACadSharp.Entities.TextEntity;
using AcadMText = ACadSharp.Entities.MText;
using System.Text;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class DwgConvertCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View view = doc.ActiveView;

            List<ImportInstance> dwgs = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(ImportInstance))
                .Cast<ImportInstance>()
                .ToList();

            if (dwgs.Count == 0)
            {
                TaskDialog.Show("HMV Tools", "No DWG imports found in the current view.");
                return Result.Cancelled;
            }

            List<string> names = new List<string>();
            foreach (var d in dwgs)
            {
                string name = d.Category != null ? d.Category.Name : "Unknown";
                names.Add(name + " [ID: " + d.Id.IntegerValue + "]");
            }

            DwgConvertWindow win = new DwgConvertWindow(names);
            if (win.ShowDialog() != true)
                return Result.Cancelled;

            List<ImportInstance> selectedDwgs = new List<ImportInstance>();
            foreach (int idx in win.SelectedIndices)
                selectedDwgs.Add(dwgs[idx]);

            string font = win.SelectedFont;

            if (win.Action == DwgConvertAction.LinesAndTexts)
            {
                ConvertLines(doc, view, selectedDwgs);
                StandardizeTexts(doc, view, selectedDwgs, font);
                return Result.Succeeded;
            }
            else if (win.Action == DwgConvertAction.FilledRegions)
            {
                return ProcessFilledRegions(doc, view, selectedDwgs);
            }

            return Result.Cancelled;
        }

        // ====================================================================
        // OPUS STABLE CODE (LINES AND TEXTS) - UNTOUCHED
        // ====================================================================

        private Result ConvertLines(Document doc, View view, List<ImportInstance> selectedDwgs)
        {
            int count = 0;
            Category linesCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
            var lineStyleCache = new Dictionary<string, GraphicsStyle>();

            using (Transaction t = new Transaction(doc, "DWG to Standardized Lines"))
            {
                t.Start();

                foreach (Category subCat in linesCat.SubCategories)
                {
                    if (subCat.Name.StartsWith("HMV_LINEA"))
                    {
                        GraphicsStyle gs = subCat.GetGraphicsStyle(GraphicsStyleType.Projection);
                        if (gs != null) lineStyleCache[subCat.Name] = gs;
                    }
                }

                foreach (ImportInstance selected in selectedDwgs)
                {
                    Options geoOptions = new Options();
                    geoOptions.View = view;
                    geoOptions.ComputeReferences = true;
                    GeometryElement geoElem = selected.get_Geometry(geoOptions);

                    var curveData = new List<(Curve curve, GraphicsStyle style)>();
                    GetCurves(doc, geoElem, curveData);

                    foreach (var item in curveData)
                    {
                        try
                        {
                            if (item.curve.Length < 0.003) continue;

                            DetailCurve dc = doc.Create.NewDetailCurve(view, item.curve);
                            string styleName = GetHmvLineStyleName(doc, item.style);

                            if (styleName != null)
                            {
                                if (!lineStyleCache.TryGetValue(styleName, out GraphicsStyle hmvStyle))
                                {
                                    hmvStyle = CreateLineStyle(doc, linesCat, item.style, styleName);
                                    if (hmvStyle != null) lineStyleCache[styleName] = hmvStyle;
                                }
                                if (hmvStyle != null) dc.LineStyle = hmvStyle;
                            }
                            count++;
                        }
                        catch { }
                    }
                }
                t.Commit();
            }

            TaskDialog.Show("HMV Tools", count + " detail lines created.\n" + lineStyleCache.Count + " HMV line styles used.");
            return Result.Succeeded;
        }

        private static readonly double[] AllowedSizes = { 1.5, 2.0, 2.5, 3.0, 3.5 };

        private Result StandardizeTexts(Document doc, View view, List<ImportInstance> selectedDwgs, string font)
        {
            int textCount = 0;
            int purgedTextTypes = 0;
            int errorCount = 0;
            List<string> errorMessages = new List<string>();
            var typeCache = new Dictionary<string, TextNoteType>();

            List<DwgTextData> allTexts = new List<DwgTextData>();

            foreach (ImportInstance dwg in selectedDwgs)
            {
                string dwgPath = GetDwgFilePath(doc, dwg);

                if (dwgPath == null || !System.IO.File.Exists(dwgPath))
                {
                    TaskDialog.Show("HMV Tools", "Cannot find DWG file on disk.\nAction cancelled for this DWG.");
                    continue;
                }

                BoundingBoxXYZ revitBB = dwg.get_BoundingBox(view);
                if (revitBB == null) continue;

                try
                {
                    CadDocument cadDoc;
                    using (DwgReader reader = new DwgReader(dwgPath))
                    {
                        cadDoc = reader.Read();
                        // DEBUG - count all entity types
                        var typeCounts = new Dictionary<string, int>();
                        void CountEntities(IEnumerable<ACadSharp.Entities.Entity> ents, string prefix)
                        {
                            foreach (var e in ents)
                            {
                                string key = prefix + e.GetType().Name;
                                if (!typeCounts.ContainsKey(key)) typeCounts[key] = 0;
                                typeCounts[key]++;

                                if (e is ACadSharp.Entities.Insert blk && blk.Block != null)
                                    CountEntities(blk.Block.Entities, prefix + "  ");

                                // Check for attributes
                                if (e is ACadSharp.Entities.Insert ins2)
                                {
                                    try
                                    {
                                        var atts = ins2.Attributes;
                                        if (atts != null && atts.Count > 0)
                                        {
                                            if (!typeCounts.ContainsKey(prefix + "ATTRIB"))
                                                typeCounts[prefix + "ATTRIB"] = 0;
                                            typeCounts[prefix + "ATTRIB"] += atts.Count;
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                        CountEntities(cadDoc.Entities, "");

                        StringBuilder debugMsg = new StringBuilder();
                        debugMsg.AppendLine("Entity types found:");
                        foreach (var kv in typeCounts.OrderBy(k => k.Key))
                            debugMsg.AppendLine(kv.Key + ": " + kv.Value);
                        TaskDialog.Show("DEBUG Entities", debugMsg.ToString());
                    }

                    List<DwgTextData> rawTexts = new List<DwgTextData>();

                    foreach (var entity in cadDoc.Entities)
                        CollectRawText(entity, null, rawTexts);

                    if (rawTexts.Count == 0) continue;

                    // ===== CALIBRATION USING LONGEST LINE =====
                    Options curveOpts = new Options();
                    curveOpts.View = view;
                    GeometryElement curveGeo = dwg.get_Geometry(curveOpts);

                    XYZ revLineMid = null;
                    double revLineAngle = 0;
                    double revLineLen = 0;
                    FindLongestRevitLine(curveGeo,
                        ref revLineMid, ref revLineAngle, ref revLineLen);

                    if (revLineMid == null)
                    {
                        TaskDialog.Show("HMV Tools", "Could not find any line in Revit geometry.");
                        continue;
                    }

                    double cadMidX = 0, cadMidY = 0;
                    double cadLineAngle = 0;
                    double cadLineLen = 0;
                    foreach (var entity in cadDoc.Entities)
                    {
                        FindLongestCadLine(entity, null,
                            ref cadMidX, ref cadMidY,
                            ref cadLineAngle, ref cadLineLen);
                    }

                    if (cadLineLen < 0.001)
                    {
                        TaskDialog.Show("HMV Tools", "Could not find any line in CAD file.");
                        continue;
                    }

                    double cadMidXft = cadMidX / 0.3048;
                    double cadMidYft = cadMidY / 0.3048;

                    double rotDelta = revLineAngle - cadLineAngle;
                    double cosR = Math.Cos(rotDelta);
                    double sinR = Math.Sin(rotDelta);

                    foreach (var rt in rawTexts)
                    {
                        double txFt = rt.LocalPosition.X / 0.3048;
                        double tyFt = rt.LocalPosition.Y / 0.3048;

                        double dx = txFt - cadMidXft;
                        double dy = tyFt - cadMidYft;

                        double rotX = dx * cosR - dy * sinR;
                        double rotY = dx * sinR + dy * cosR;

                        XYZ finalPos = new XYZ(
                            revLineMid.X + rotX,
                            revLineMid.Y + rotY,
                            0);

                        double finalRotation = rt.Rotation + rotDelta;
                        double heightMm = rt.Height * 1000.0;

                        allTexts.Add(new DwgTextData(
                            rt.Text, finalPos, heightMm, finalRotation,
                            rt.HAlign, rt.VAlignParam));
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("HMV Tools", "Error reading DWG: " + ex.Message);
                    continue;
                }
            }

            if (allTexts.Count == 0) return Result.Cancelled;

            using (Transaction t = new Transaction(doc, "DWG Texts to Standardized Notes"))
            {
                t.Start();
                FailureHandlingOptions failureOptions = t.GetFailureHandlingOptions();
                t.SetFailureHandlingOptions(failureOptions);

                foreach (TextNoteType tnt in new FilteredElementCollector(doc).OfClass(typeof(TextNoteType)).Cast<TextNoteType>())
                {
                    if (tnt.Name.StartsWith("HMV_General_")) typeCache[tnt.Name] = tnt;
                }

                TextNoteType baseType = new FilteredElementCollector(doc).OfClass(typeof(TextNoteType)).Cast<TextNoteType>().FirstOrDefault();

                if (baseType == null)
                {
                    t.RollBack();
                    return Result.Cancelled;
                }

                foreach (DwgTextData td in allTexts)
                {
                    try
                    {
                        double snapped = SnapToAllowed(td.Height);
                        string typeName = "HMV_General_" + snapped.ToString("0.0") + " " + font;

                        if (!typeCache.TryGetValue(typeName, out TextNoteType hmvType))
                        {
                            hmvType = CreateTextType(doc, baseType, typeName, snapped, font);
                            if (hmvType != null) typeCache[typeName] = hmvType;
                        }

                        if (hmvType != null)
                        {
                            try
                            {
                                TextNoteOptions options = new TextNoteOptions(hmvType.Id);
                                options.HorizontalAlignment = td.HAlign;
                                options.Rotation = td.Rotation;
                                TextNote note = TextNote.Create(doc, view.Id, td.LocalPosition, td.Text, options);
                                Parameter vAlignParam = note.get_Parameter(BuiltInParameter.TEXT_ALIGN_VERT);
                                if (vAlignParam != null && !vAlignParam.IsReadOnly)
                                    vAlignParam.Set(td.VAlignParam);

                                textCount++;
                            }
                            catch (Exception)
                            {
                                try
                                {
                                    TextNote note = TextNote.Create(doc, view.Id, td.LocalPosition, td.Text, hmvType.Id);
                                    textCount++;
                                }
                                catch (Exception ex2)
                                {
                                    errorCount++;
                                    if (errorMessages.Count < 5) errorMessages.Add($"[{td.Text}] : {ex2.Message}");
                                }
                            }
                        }
                    }
                    catch { }
                }

                HashSet<ElementId> usedTextTypeIds = new HashSet<ElementId>(new FilteredElementCollector(doc).OfClass(typeof(TextNote)).Cast<TextNote>().Select(tn => tn.GetTypeId()));
                foreach (TextNoteType tnt in new FilteredElementCollector(doc).OfClass(typeof(TextNoteType)).Cast<TextNoteType>().ToList())
                {
                    if (tnt.Name.StartsWith("HMV_General_") && !usedTextTypeIds.Contains(tnt.Id))
                    {
                        try { doc.Delete(tnt.Id); purgedTextTypes++; } catch { }
                    }
                }
                t.Commit();
            }

            StringBuilder resultMsg = new StringBuilder();
            resultMsg.AppendLine($"Texts created: {textCount}");
            resultMsg.AppendLine($"Cleanup:\nUnused text types purged: {purgedTextTypes}");

            if (errorCount > 0)
            {
                resultMsg.AppendLine("\n--- DEBUG INFO ---");
                resultMsg.AppendLine($"Failed to create {errorCount} texts. Top errors:");
                foreach (string err in errorMessages) resultMsg.AppendLine($"- {err}");
            }

            TaskDialog.Show("HMV Tools", resultMsg.ToString());
            return Result.Succeeded;
        }


        // ====================================================================
        // NEW LOGIC: FILLED REGIONS (Re-using Opus Math)
        // ====================================================================

        // ====================================================================
        // NEW LOGIC: FILLED REGIONS (STRUCTURE ANALYZER ONLY)
        // ====================================================================

        private Result ProcessFilledRegions(Document doc, View view, List<ImportInstance> selectedDwgs)
        {
            using (Transaction t = new Transaction(doc, "DWG Filled Regions Analyzer"))
            {
                t.Start();

                foreach (ImportInstance dwg in selectedDwgs)
                {
                    string dwgPath = GetDwgFilePath(doc, dwg);
                    if (dwgPath == null || !System.IO.File.Exists(dwgPath)) continue;

                    try
                    {
                        CadDocument cadDoc;
                        using (DwgReader reader = new DwgReader(dwgPath)) cadDoc = reader.Read();

                        List<DwgHatchData> rawHatches = new List<DwgHatchData>();
                        foreach (var entity in cadDoc.Entities)
                        {
                            CollectRawHatchWithTransform(entity, 0, 0, 1, 0, rawHatches);
                        }

                        System.Text.StringBuilder analyzerReport = new System.Text.StringBuilder();
                        analyzerReport.AppendLine($"Total hatches found in CAD: {rawHatches.Count}");

                        if (rawHatches.Count > 0)
                        {
                            try
                            {
                                // Analyze the very first hatch to see what properties it actually has
                                object hatchObj = rawHatches[0].Hatch;
                                Type hatchType = hatchObj.GetType();

                                analyzerReport.AppendLine("\n=== HATCH PROPERTIES ===");
                                foreach (var prop in hatchType.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
                                {
                                    analyzerReport.AppendLine($"- {prop.Name} ({prop.PropertyType.Name})");
                                }

                                dynamic dynHatch = hatchObj;
                                if (dynHatch.Paths != null)
                                {
                                    analyzerReport.AppendLine("\n=== PATH PROPERTIES ===");
                                    foreach (dynamic path in dynHatch.Paths)
                                    {
                                        Type pathType = path.GetType();
                                        foreach (var prop in pathType.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
                                        {
                                            analyzerReport.AppendLine($"- {prop.Name} ({prop.PropertyType.Name})");
                                        }
                                        break; // We only need to look at the first path
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                analyzerReport.AppendLine($"Analyzer Error: {ex.Message}");
                            }
                        }

                        TaskDialog debugInfo = new TaskDialog("Hatch Structure Analyzer");
                        debugInfo.MainInstruction = "Please copy this list for me";
                        debugInfo.MainContent = analyzerReport.ToString();
                        debugInfo.Show();

                        // We just need the report, so we cancel the transaction
                        t.RollBack();
                        return Result.Cancelled;
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("HMV Tools", "Error reading DWG Regions: " + ex.Message);
                    }
                }

                t.RollBack();
            }
            return Result.Cancelled;
        }
        // Applies block transform, then Opus calibration transform
        private XYZ ApplyHatchTransform(double x, double y, DwgHatchData hd, double cadMidXft, double cadMidYft, double cosR, double sinR, XYZ revLineMid)
        {
            // 1. Block
            double bx = x * hd.Scale * Math.Cos(hd.Rot) - y * hd.Scale * Math.Sin(hd.Rot) + hd.TX;
            double by = x * hd.Scale * Math.Sin(hd.Rot) + y * hd.Scale * Math.Cos(hd.Rot) + hd.TY;

            // 2. Opus Math
            double txFt = bx / 0.3048;
            double tyFt = by / 0.3048;
            double dx = txFt - cadMidXft;
            double dy = tyFt - cadMidYft;

            double rotX = dx * cosR - dy * sinR;
            double rotY = dx * sinR + dy * cosR;

            return new XYZ(revLineMid.X + rotX, revLineMid.Y + rotY, 0);
        }

        private class DwgHatchData
        {
            public ACadSharp.Entities.Hatch Hatch;
            public double TX, TY, Scale, Rot;
            public DwgHatchData(ACadSharp.Entities.Hatch h, double tx, double ty, double s, double r)
            {
                Hatch = h; TX = tx; TY = ty; Scale = s; Rot = r;
            }
        }

        private void CollectRawHatchWithTransform(ACadSharp.Entities.Entity entity, double tx, double ty, double tscale, double trot, List<DwgHatchData> rawHatches)
        {
            if (entity is ACadSharp.Entities.Insert nested && nested.Block != null)
            {
                double nx = nested.InsertPoint.X;
                double ny = nested.InsertPoint.Y;
                double newX = nx * tscale * Math.Cos(trot) - ny * tscale * Math.Sin(trot) + tx;
                double newY = nx * tscale * Math.Sin(trot) + ny * tscale * Math.Cos(trot) + ty;
                foreach (var be in nested.Block.Entities)
                {
                    CollectRawHatchWithTransform(be, newX, newY, tscale * nested.XScale, trot + nested.Rotation, rawHatches);
                }
            }
            else if (entity is ACadSharp.Entities.Hatch hatch)
            {
                rawHatches.Add(new DwgHatchData(hatch, tx, ty, tscale, trot));
            }
        }

        private List<Curve> SortAndCloseCurves(List<Curve> curves)
        {
            if (curves.Count < 2) return curves;
            List<Curve> sorted = new List<Curve>();
            sorted.Add(curves[0]);
            curves.RemoveAt(0);

            while (curves.Count > 0)
            {
                XYZ endPt = sorted.Last().GetEndPoint(1);
                bool found = false;
                for (int i = 0; i < curves.Count; i++)
                {
                    if (curves[i].GetEndPoint(0).DistanceTo(endPt) < 0.01)
                    {
                        sorted.Add(curves[i]);
                        curves.RemoveAt(i);
                        found = true; break;
                    }
                    if (curves[i].GetEndPoint(1).DistanceTo(endPt) < 0.01)
                    {
                        sorted.Add(curves[i].CreateReversed());
                        curves.RemoveAt(i);
                        found = true; break;
                    }
                }
                if (!found) break;
            }

            if (sorted.Count > 1)
            {
                XYZ firstPt = sorted.First().GetEndPoint(0);
                XYZ lastPt = sorted.Last().GetEndPoint(1);
                if (firstPt.DistanceTo(lastPt) > 0.001)
                {
                    try { sorted.Add(Line.CreateBound(lastPt, firstPt)); } catch { }
                }
            }
            return sorted;
        }


        // ====================================================================
        // OPUS SUPPORT FUNCTIONS - UNTOUCHED
        // ====================================================================

        private void FindLongestRevitLine(GeometryElement geoElem,
            ref XYZ midpoint, ref double angle, ref double maxLen)
        {
            foreach (GeometryObject geoObj in geoElem)
            {
                if (geoObj is Autodesk.Revit.DB.Line line)
                {
                    try
                    {
                        double len = line.Length;
                        if (len > maxLen)
                        {
                            maxLen = len;
                            XYZ p1 = line.GetEndPoint(0);
                            XYZ p2 = line.GetEndPoint(1);
                            midpoint = (p1 + p2) / 2.0;
                            angle = Math.Atan2(p2.Y - p1.Y, p2.X - p1.X);
                        }
                    }
                    catch { }
                }
                else if (geoObj is PolyLine pl)
                {
                    IList<XYZ> pts = pl.GetCoordinates();
                    for (int i = 0; i < pts.Count - 1; i++)
                    {
                        XYZ p1 = pts[i];
                        XYZ p2 = pts[i + 1];
                        double len = p1.DistanceTo(p2);
                        if (len > maxLen)
                        {
                            maxLen = len;
                            midpoint = (p1 + p2) / 2.0;
                            angle = Math.Atan2(p2.Y - p1.Y, p2.X - p1.X);
                        }
                    }
                }
                else if (geoObj is GeometryInstance gi)
                {
                    FindLongestRevitLine(gi.GetInstanceGeometry(),
                        ref midpoint, ref angle, ref maxLen);
                }
                else if (geoObj is GeometryElement nested)
                {
                    FindLongestRevitLine(nested,
                        ref midpoint, ref angle, ref maxLen);
                }
            }
        }

        private void FindLongestCadLine(ACadSharp.Entities.Entity entity,
            ACadSharp.Entities.Insert parentInsert,
            ref double midX, ref double midY,
            ref double angle, ref double maxLen)
        {
            if (entity is ACadSharp.Entities.Line ln)
            {
                double x1 = ln.StartPoint.X;
                double y1 = ln.StartPoint.Y;
                double x2 = ln.EndPoint.X;
                double y2 = ln.EndPoint.Y;

                if (parentInsert != null)
                {
                    double ix = parentInsert.InsertPoint.X;
                    double iy = parentInsert.InsertPoint.Y;
                    double sc = parentInsert.XScale;
                    double ir = parentInsert.Rotation;

                    double nx1 = x1 * sc * Math.Cos(ir) - y1 * sc * Math.Sin(ir) + ix;
                    double ny1 = x1 * sc * Math.Sin(ir) + y1 * sc * Math.Cos(ir) + iy;
                    double nx2 = x2 * sc * Math.Cos(ir) - y2 * sc * Math.Sin(ir) + ix;
                    double ny2 = x2 * sc * Math.Sin(ir) + y2 * sc * Math.Cos(ir) + iy;
                    x1 = nx1; y1 = ny1; x2 = nx2; y2 = ny2;
                }

                double dx = x2 - x1;
                double dy = y2 - y1;
                double len = Math.Sqrt(dx * dx + dy * dy);

                if (len > maxLen)
                {
                    maxLen = len;
                    midX = (x1 + x2) / 2.0;
                    midY = (y1 + y2) / 2.0;
                    angle = Math.Atan2(dy, dx);
                }
            }
            else if (entity is ACadSharp.Entities.Insert ins && ins.Block != null)
            {
                foreach (var be in ins.Block.Entities)
                {
                    FindLongestCadLine(be, ins,
                        ref midX, ref midY, ref angle, ref maxLen);
                }
            }
        }

        private string GetDwgFilePath(Document doc, ImportInstance import)
        {
            try
            {
                Element typeElem = doc.GetElement(import.GetTypeId());
                if (typeElem is CADLinkType linkType && linkType.IsExternalFileReference())
                {
                    ExternalFileReference extRef = linkType.GetExternalFileReference();
                    if (extRef != null)
                    {
                        string path = ModelPathUtils.ConvertModelPathToUserVisiblePath(extRef.GetPath());
                        if (System.IO.File.Exists(path)) return path;
                    }
                }
            }
            catch { }

            string fileName = import.Category?.Name ?? "DWG file";
            if (!fileName.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase)) fileName += ".dwg";

            try
            {
                if (!string.IsNullOrEmpty(doc.PathName))
                {
                    string dir = System.IO.Path.GetDirectoryName(doc.PathName);
                    string localPath = System.IO.Path.Combine(dir, fileName);
                    if (System.IO.File.Exists(localPath)) return localPath;
                }
            }
            catch { }

            TaskDialog.Show("HMV Tools: Missing Link",
                $"The file '{fileName}' is IMPORTED, not linked, so its path is missing.\n\n" +
                "In the next window, please locate the original DWG file on your computer so we can extract its texts.");

            try
            {
                var dialog = new Microsoft.Win32.OpenFileDialog();
                dialog.Title = "Locate original DWG: " + fileName;
                dialog.Filter = "DWG files (*.dwg)|*.dwg";

                if (dialog.ShowDialog() == true) return dialog.FileName;
            }
            catch { }

            return null;
        }

        private void CollectRawText(ACadSharp.Entities.Entity entity,
             ACadSharp.Entities.Insert parentInsert,
             List<DwgTextData> rawTexts)
        {
            if (entity is ACadSharp.Entities.Insert nestedInsert
                && nestedInsert.Block != null)
            {
                ACadSharp.Entities.Insert effectiveInsert = nestedInsert;
                if (parentInsert != null)
                {
                    double px = parentInsert.InsertPoint.X;
                    double py = parentInsert.InsertPoint.Y;
                    double psc = parentInsert.XScale;
                    double pr = parentInsert.Rotation;

                    double nx = nestedInsert.InsertPoint.X;
                    double ny = nestedInsert.InsertPoint.Y;

                    double newX = nx * psc * Math.Cos(pr)
                                - ny * psc * Math.Sin(pr) + px;
                    double newY = nx * psc * Math.Sin(pr)
                                + ny * psc * Math.Cos(pr) + py;

                    foreach (var be in nestedInsert.Block.Entities)
                    {
                        CollectRawTextWithTransform(be,
                            newX, newY,
                            psc * nestedInsert.XScale,
                            pr + nestedInsert.Rotation,
                            rawTexts);
                    }
                }
                else
                {
                    foreach (var be in nestedInsert.Block.Entities)
                    {
                        CollectRawText(be, nestedInsert, rawTexts);
                    }
                }
                return;
            }

            CollectRawTextWithTransform(entity,
                parentInsert?.InsertPoint.X ?? 0,
                parentInsert?.InsertPoint.Y ?? 0,
                parentInsert?.XScale ?? 1.0,
                parentInsert?.Rotation ?? 0,
                rawTexts);
        }

        private void CollectRawTextWithTransform(
            ACadSharp.Entities.Entity entity,
            double tx, double ty, double tscale, double trot,
            List<DwgTextData> rawTexts)
        {
            if (entity is ACadSharp.Entities.Insert nested
                && nested.Block != null)
            {
                double nx = nested.InsertPoint.X;
                double ny = nested.InsertPoint.Y;

                double newX = nx * tscale * Math.Cos(trot)
                            - ny * tscale * Math.Sin(trot) + tx;
                double newY = nx * tscale * Math.Sin(trot)
                            + ny * tscale * Math.Cos(trot) + ty;

                foreach (var be in nested.Block.Entities)
                {
                    CollectRawTextWithTransform(be,
                        newX, newY,
                        tscale * nested.XScale,
                        trot + nested.Rotation,
                        rawTexts);
                }
                return;
            }

            string text = null;
            double x = 0, y = 0, height = 0, rotation = 0;
            HorizontalTextAlignment hAlign = HorizontalTextAlignment.Left;
            int vAlign = 3;

            if (entity is AcadText textEnt)
            {
                text = textEnt.Value;
                height = textEnt.Height;
                rotation = textEnt.Rotation;

                if (textEnt.HorizontalAlignment != ACadSharp.Entities.TextHorizontalAlignment.Left ||
                    textEnt.VerticalAlignment != ACadSharp.Entities.TextVerticalAlignmentType.Baseline)
                {
                    x = textEnt.AlignmentPoint.X;
                    y = textEnt.AlignmentPoint.Y;
                }
                else
                {
                    x = textEnt.InsertPoint.X;
                    y = textEnt.InsertPoint.Y;
                }

                if (textEnt.HorizontalAlignment == ACadSharp.Entities.TextHorizontalAlignment.Center
                    || textEnt.HorizontalAlignment == ACadSharp.Entities.TextHorizontalAlignment.Middle)
                    hAlign = HorizontalTextAlignment.Center;
                else if (textEnt.HorizontalAlignment == ACadSharp.Entities.TextHorizontalAlignment.Right)
                    hAlign = HorizontalTextAlignment.Right;

                if (textEnt.VerticalAlignment == ACadSharp.Entities.TextVerticalAlignmentType.Top) vAlign = 1;
                else if (textEnt.VerticalAlignment == ACadSharp.Entities.TextVerticalAlignmentType.Middle) vAlign = 2;
            }
            else if (entity is AcadMText mtext)
            {
                text = StripMTextFormatting(mtext.Value);
                x = mtext.InsertPoint.X;
                y = mtext.InsertPoint.Y;
                height = mtext.Height;
                rotation = mtext.Rotation;

                int attach = (int)mtext.AttachmentPoint;
                if (attach == 2 || attach == 5 || attach == 8) hAlign = HorizontalTextAlignment.Center;
                else if (attach == 3 || attach == 6 || attach == 9) hAlign = HorizontalTextAlignment.Right;

                if (attach >= 1 && attach <= 3) vAlign = 1;
                else if (attach >= 4 && attach <= 6) vAlign = 2;
            }
            // Check for attributed inserts (block attributes)
            if (entity is ACadSharp.Entities.Insert attIns)
            {
                try
                {
                    foreach (var att in attIns.Attributes)
                    {
                        string attText = att.Value;
                        if (string.IsNullOrWhiteSpace(attText)) continue;

                        double ax = att.InsertPoint.X;
                        double ay = att.InsertPoint.Y;
                        double aHeight = att.Height;
                        double aRot = att.Rotation;

                        // Apply accumulated transform
                        double arx = ax * tscale * Math.Cos(trot)
                                  - ay * tscale * Math.Sin(trot) + tx;
                        double ary = ax * tscale * Math.Sin(trot)
                                  + ay * tscale * Math.Cos(trot) + ty;

                        aHeight *= tscale;
                        aRot += trot;

                        rawTexts.Add(new DwgTextData(attText,
                            new XYZ(arx, ary, 0), aHeight, aRot,
                            HorizontalTextAlignment.Left, 3));
                    }
                }
                catch { }
            }

            if (string.IsNullOrWhiteSpace(text)) return;

            double rx = x * tscale * Math.Cos(trot)
                      - y * tscale * Math.Sin(trot) + tx;
            double ry = x * tscale * Math.Sin(trot)
                      + y * tscale * Math.Cos(trot) + ty;

            height *= tscale;
            rotation += trot;

            rawTexts.Add(new DwgTextData(text,
                new XYZ(rx, ry, 0), height, rotation, hAlign, vAlign));
        }

        private string StripMTextFormatting(string mtext)
        {
            if (string.IsNullOrEmpty(mtext)) return mtext;
            string result = mtext;
            result = System.Text.RegularExpressions.Regex.Replace(result, @"\{\\[^;]+;([^}]*)}", "$1");
            result = System.Text.RegularExpressions.Regex.Replace(result, @"\\A\d;", "");
            result = result.Replace("\\P", "\n");
            result = System.Text.RegularExpressions.Regex.Replace(result, @"\\[a-zA-Z][^;]*;", "");
            return result.Replace("{", "").Replace("}", "").Trim();
        }

        private class DwgTextData
        {
            public string Text;
            public XYZ LocalPosition;
            public double Height;
            public double Rotation;
            public HorizontalTextAlignment HAlign;
            public int VAlignParam;

            public DwgTextData(string text, XYZ pos, double height, double rot, HorizontalTextAlignment hAlign, int vAlign)
            {
                Text = text;
                LocalPosition = pos;
                Height = height;
                Rotation = rot;
                HAlign = hAlign;
                VAlignParam = vAlign;
            }
        }

        private void GetCurves(Document doc, GeometryElement geoElem, List<(Curve curve, GraphicsStyle style)> curveData)
        {
            foreach (GeometryObject geoObj in geoElem)
            {
                GraphicsStyle gs = doc.GetElement(geoObj.GraphicsStyleId) as GraphicsStyle;
                if (geoObj is Curve curve) curveData.Add((curve, gs));
                else if (geoObj is PolyLine polyline)
                {
                    IList<XYZ> pts = polyline.GetCoordinates();
                    for (int i = 0; i < pts.Count - 1; i++)
                    {
                        if (pts[i].DistanceTo(pts[i + 1]) < 0.003) continue;
                        curveData.Add((Autodesk.Revit.DB.Line.CreateBound(pts[i], pts[i + 1]), gs));
                    }
                }
                else if (geoObj is GeometryInstance geoInst) GetCurves(doc, geoInst.GetInstanceGeometry(), curveData);
                else if (geoObj is GeometryElement nestedElem) GetCurves(doc, nestedElem, curveData);
            }
        }

        private string GetHmvLineStyleName(Document doc, GraphicsStyle style)
        {
            if (style == null || style.GraphicsStyleCategory == null) return null;
            int w = style.GraphicsStyleCategory.GetLineWeight(GraphicsStyleType.Projection) ?? 1;
            ElementId patternId = style.GraphicsStyleCategory.GetLinePatternId(GraphicsStyleType.Projection);
            string patternName = (doc.GetElement(patternId) as LinePatternElement)?.Name.ToUpper() ?? "SOLID";
            return "HMV_LINEA " + patternName + " " + w;
        }

        private GraphicsStyle CreateLineStyle(Document doc, Category linesCat, GraphicsStyle sourceStyle, string name)
        {
            try
            {
                Category newCat = doc.Settings.Categories.NewSubcategory(linesCat, name);
                Category srcCat = sourceStyle.GraphicsStyleCategory;

                int? weight = srcCat.GetLineWeight(GraphicsStyleType.Projection);
                if (weight.HasValue) newCat.SetLineWeight(weight.Value, GraphicsStyleType.Projection);

                ElementId patternId = srcCat.GetLinePatternId(GraphicsStyleType.Projection);
                if (patternId != null && patternId != ElementId.InvalidElementId)
                    newCat.SetLinePatternId(patternId, GraphicsStyleType.Projection);

                newCat.LineColor = new Autodesk.Revit.DB.Color(0, 0, 0);
                return newCat.GetGraphicsStyle(GraphicsStyleType.Projection);
            }
            catch { return null; }
        }

        private double SnapToAllowed(double sizeMm)
        {
            double closest = AllowedSizes[0];
            double minDiff = Math.Abs(sizeMm - closest);
            foreach (double a in AllowedSizes)
            {
                double diff = Math.Abs(sizeMm - a);
                if (diff < minDiff) { minDiff = diff; closest = a; }
            }
            return closest;
        }

        private TextNoteType CreateTextType(Document doc, TextNoteType baseType, string name, double sizeMm, string font)
        {
            try
            {
                TextNoteType nt = baseType.Duplicate(name) as TextNoteType;
                nt.get_Parameter(BuiltInParameter.TEXT_SIZE).Set(sizeMm / 304.8);
                nt.get_Parameter(BuiltInParameter.TEXT_FONT).Set(font);
                nt.get_Parameter(BuiltInParameter.TEXT_WIDTH_SCALE).Set(1.0);
                nt.get_Parameter(BuiltInParameter.TEXT_BACKGROUND).Set(1);
                return nt;
            }
            catch { return null; }
        }
    }
}