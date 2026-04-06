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

            if (win.Action == DwgConvertAction.ConvertLines)
                return ConvertLines(doc, view, selectedDwgs);
            else if (win.Action == DwgConvertAction.StandardizeTexts)
                return StandardizeTexts(doc, view, selectedDwgs, font);

            return Result.Cancelled;
        }

        // ========== LINES ==========
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

            TaskDialog.Show("HMV Tools", count + " detail lines created.\n" + lineStyleCache.Count + " HMV line styles used.\n\nNext: Select DWG(s) again and click 'Standardize Texts'.");
            return Result.Succeeded;
        }

        // ========== TEXTS ==========
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
                    }

                    List<DwgTextData> rawTexts = new List<DwgTextData>();

                    foreach (var entity in cadDoc.Entities)
                        CollectRawText(entity, null, rawTexts);

                    foreach (var entity in cadDoc.Entities)
                    {
                        if (entity is ACadSharp.Entities.Insert insert && insert.Block != null)
                        {
                            foreach (var be in insert.Block.Entities)
                                CollectRawText(be, insert, rawTexts);
                        }
                    }

                    if (rawTexts.Count == 0) continue;

                    // --- RIGID MATH (RESTORED FROM REFERENCE + FIX FOR ROTATION) ---

                    // 1. Get Revit Center & Exact Rotation
                    double revitCenterX = (revitBB.Min.X + revitBB.Max.X) / 2.0;
                    double revitCenterY = (revitBB.Min.Y + revitBB.Max.Y) / 2.0;

                    Transform dwgTransform = dwg.GetTransform();
                    double revitRot = Math.Atan2(dwgTransform.BasisX.Y, dwgTransform.BasisX.X);

                    // 2. Convert CAD units to Feet and find CAD Average Center 
                    // (This prevents independent X/Y stretching and preserves perfect relative distances)
                    double sumX = 0, sumY = 0;
                    foreach (var rt in rawTexts)
                    {
                        // Match Reference Script rigid scaling (DWG Meters to Revit Feet)
                        double feetX = rt.LocalPosition.X / 0.3048;
                        double feetY = rt.LocalPosition.Y / 0.3048;
                        rt.LocalPosition = new XYZ(feetX, feetY, 0);

                        sumX += feetX;
                        sumY += feetY;
                    }
                    double cadCenterX = sumX / rawTexts.Count;
                    double cadCenterY = sumY / rawTexts.Count;

                    // 3. Map points: Translate to origin -> Rotate globally -> Translate to Revit Center
                    foreach (var rt in rawTexts)
                    {
                        // Rigid distance from CAD average center
                        double dx = rt.LocalPosition.X - cadCenterX;
                        double dy = rt.LocalPosition.Y - cadCenterY;

                        // Rotate the entire block mathematically
                        double rotX = dx * Math.Cos(revitRot) - dy * Math.Sin(revitRot);
                        double rotY = dx * Math.Sin(revitRot) + dy * Math.Cos(revitRot);

                        // Position it at the Revit bounding box center
                        XYZ finalPos = new XYZ(revitCenterX + rotX, revitCenterY + rotY, 0);

                        // The individual text rotation is its CAD rotation + the Revit global rotation
                        double finalRotation = rt.Rotation + revitRot;

                        // Height parameter matching reference script (Meters to MM)
                        double heightMm = rt.Height * 1000.0;

                        allTexts.Add(new DwgTextData(rt.Text, finalPos, heightMm, finalRotation, rt.HAlign, rt.VAlignParam));
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("HMV Tools", "Error reading DWG file structures: " + ex.Message);
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

                // Cleanup unused types
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

        // ===== DWG FILE FINDER =====
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

        // ===== ACadSharp TEXT EXTRACTION =====
        private void CollectRawText(ACadSharp.Entities.Entity entity, ACadSharp.Entities.Insert parentInsert, List<DwgTextData> rawTexts)
        {
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

                if (textEnt.HorizontalAlignment == ACadSharp.Entities.TextHorizontalAlignment.Center || textEnt.HorizontalAlignment == ACadSharp.Entities.TextHorizontalAlignment.Middle) hAlign = HorizontalTextAlignment.Center;
                else if (textEnt.HorizontalAlignment == ACadSharp.Entities.TextHorizontalAlignment.Right) hAlign = HorizontalTextAlignment.Right;

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

            if (string.IsNullOrWhiteSpace(text)) return;

            if (parentInsert != null)
            {
                double ix = parentInsert.InsertPoint.X;
                double iy = parentInsert.InsertPoint.Y;
                double scale = parentInsert.XScale;
                double insRot = parentInsert.Rotation;

                double rx = x * scale * Math.Cos(insRot) - y * scale * Math.Sin(insRot) + ix;
                double ry = x * scale * Math.Sin(insRot) + y * scale * Math.Cos(insRot) + iy;

                x = rx;
                y = ry;
                height *= scale;
                rotation += insRot;
            }

            rawTexts.Add(new DwgTextData(text, new XYZ(x, y, 0), height, rotation, hAlign, vAlign));
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

        // ========== LINE HELPERS ==========
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

    public enum DwgConvertAction { ConvertLines, StandardizeTexts }
}