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

            List<ImportInstance> dwgs = new FilteredElementCollector(doc)
                .OfClass(typeof(ImportInstance))
                .Cast<ImportInstance>()
                .ToList();

            if (dwgs.Count == 0)
            {
                TaskDialog.Show("HMV Tools", "No DWG imports found.");
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
        private Result ConvertLines(Document doc, View view,
            List<ImportInstance> selectedDwgs)
        {
            int count = 0;
            Category linesCat = doc.Settings.Categories
                .get_Item(BuiltInCategory.OST_Lines);
            var lineStyleCache = new Dictionary<string, GraphicsStyle>();

            using (Transaction t = new Transaction(doc, "DWG to Standardized Lines"))
            {
                t.Start();

                foreach (Category subCat in linesCat.SubCategories)
                {
                    if (subCat.Name.StartsWith("HMV_LINEA"))
                    {
                        GraphicsStyle gs = subCat.GetGraphicsStyle(
                            GraphicsStyleType.Projection);
                        if (gs != null)
                            lineStyleCache[subCat.Name] = gs;
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

                            DetailCurve dc = doc.Create.NewDetailCurve(
                                view, item.curve);
                            string styleName = GetHmvLineStyleName(
                                doc, item.style);

                            if (styleName != null)
                            {
                                if (!lineStyleCache.TryGetValue(styleName,
                                    out GraphicsStyle hmvStyle))
                                {
                                    hmvStyle = CreateLineStyle(doc, linesCat,
                                        item.style, styleName);
                                    if (hmvStyle != null)
                                        lineStyleCache[styleName] = hmvStyle;
                                }
                                if (hmvStyle != null)
                                    dc.LineStyle = hmvStyle;
                            }
                            count++;
                        }
                        catch { }
                    }
                }
                t.Commit();
            }

            TaskDialog.Show("HMV Tools",
                count + " detail lines created.\n"
                + lineStyleCache.Count + " HMV line styles used.\n\n"
                + "Next: Select DWG(s) again and click 'Standardize Texts'.");
            return Result.Succeeded;
        }

        // ========== TEXTS ==========
        private static readonly double[] AllowedSizes =
            { 1.5, 2.0, 2.5, 3.0, 3.5 };

        private Result StandardizeTexts(Document doc, View view,
            List<ImportInstance> selectedDwgs, string font)
        {
            int textCount = 0;
            int purgedStyles = 0;
            int purgedTextTypes = 0;
            var typeCache = new Dictionary<string, TextNoteType>();

            List<DwgTextData> allTexts = new List<DwgTextData>();

            foreach (ImportInstance dwg in selectedDwgs)
            {
                string dwgPath = GetDwgFilePath(doc, dwg);

                if (dwgPath == null || !System.IO.File.Exists(dwgPath))
                {
                    TaskDialog.Show("HMV Tools",
                        "Cannot find DWG file on disk.\n"
                        + "Make sure the DWG is linked (not imported) "
                        + "or browse to the original file.");
                    continue;
                }

                BoundingBoxXYZ revitBB = dwg.get_BoundingBox(view);
                if (revitBB == null)
                {
                    TaskDialog.Show("HMV Tools",
                        "Cannot get bounding box of DWG in current view.");
                    continue;
                }

                try
                {
                    CadDocument cadDoc;
                    using (DwgReader reader = new DwgReader(dwgPath))
                    {
                        cadDoc = reader.Read();
                    }

                    // Collect texts
                    List<double[]> rawCoords = new List<double[]>();
                    List<DwgTextData> rawTexts = new List<DwgTextData>();

                    foreach (var entity in cadDoc.Entities)
                    {
                        CollectRawText(entity, null, rawCoords, rawTexts);
                    }

                    foreach (var entity in cadDoc.Entities)
                    {
                        if (entity is ACadSharp.Entities.Insert insert
                            && insert.Block != null)
                        {
                            foreach (var be in insert.Block.Entities)
                            {
                                CollectRawText(be, insert, rawCoords, rawTexts);
                            }
                        }
                    }

                    if (rawTexts.Count == 0) continue;

                    // Get ALL entity bounds recursively including blocks
                    double allMinX = double.MaxValue, allMaxX = double.MinValue;
                    double allMinY = double.MaxValue, allMaxY = double.MinValue;

                    foreach (var entity in cadDoc.Entities)
                    {
                        CollectAllBounds(entity, cadDoc,
                            ref allMinX, ref allMinY, ref allMaxX, ref allMaxY);
                    }

                    // Fallback to text bounds if no geometry found
                    if (allMinX == double.MaxValue)
                    {
                        foreach (var rc in rawCoords)
                        {
                            if (rc[0] < allMinX) allMinX = rc[0];
                            if (rc[1] < allMinY) allMinY = rc[1];
                            if (rc[0] > allMaxX) allMaxX = rc[0];
                            if (rc[1] > allMaxY) allMaxY = rc[1];
                        }
                    }

                    double dwgWidth = allMaxX - allMinX;
                    double dwgHeight = allMaxY - allMinY;
                    double revitWidth = revitBB.Max.X - revitBB.Min.X;
                    double revitHeight = revitBB.Max.Y - revitBB.Min.Y;

                    if (dwgWidth < 0.001 || dwgHeight < 0.001) continue;

                    // Proportional mapping: DWG space -> Revit space
                    foreach (var rt in rawTexts)
                    {
                        double tX = (rt.Position.X - allMinX) / dwgWidth;
                        double tY = (rt.Position.Y - allMinY) / dwgHeight;

                        double revitX = revitBB.Min.X + tX * revitWidth;
                        double revitY = revitBB.Min.Y + tY * revitHeight;

                        XYZ newPos = new XYZ(revitX, revitY, 0);

                        allTexts.Add(new DwgTextData(
                            rt.Text, newPos, rt.HeightMm, 0));
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("HMV Tools",
                        "Error reading DWG: " + ex.Message);
                    continue;
                }
            }

            if (allTexts.Count == 0)
            {
                TaskDialog.Show("HMV Tools", "No texts found in selected DWG(s).");
                return Result.Cancelled;
            }

            using (Transaction t = new Transaction(doc,
                "DWG Texts to Standardized Notes"))
            {
                t.Start();

                foreach (TextNoteType tnt in new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>())
                {
                    if (tnt.Name.StartsWith("HMV_General_"))
                        typeCache[tnt.Name] = tnt;
                }

                TextNoteType baseType = new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .FirstOrDefault();

                if (baseType == null)
                {
                    t.RollBack();
                    TaskDialog.Show("HMV Tools", "No text types found.");
                    return Result.Cancelled;
                }

                foreach (DwgTextData td in allTexts)
                {
                    try
                    {
                        double snapped = SnapToAllowed(td.HeightMm);
                        string typeName = "HMV_General_"
                            + snapped.ToString("0.0") + " " + font;

                        if (!typeCache.TryGetValue(typeName,
                            out TextNoteType hmvType))
                        {
                            hmvType = CreateTextType(
                                doc, baseType, typeName, snapped, font);
                            if (hmvType != null)
                                typeCache[typeName] = hmvType;
                        }

                        if (hmvType != null)
                        {
                            TextNote note = TextNote.Create(
                                doc, view.Id, td.Position, td.Text, hmvType.Id);
                            textCount++;
                        }
                    }
                    catch { }
                }

                // ===== PURGE UNUSED LINE STYLES =====
                Category linesCat = doc.Settings.Categories
                    .get_Item(BuiltInCategory.OST_Lines);

                HashSet<ElementId> usedStyleIds = new HashSet<ElementId>();
                foreach (CurveElement ce in new FilteredElementCollector(doc)
                    .OfClass(typeof(CurveElement))
                    .Cast<CurveElement>())
                {
                    try
                    {
                        GraphicsStyle gs = ce.LineStyle as GraphicsStyle;
                        if (gs != null) usedStyleIds.Add(gs.Id);
                    }
                    catch { }
                }

                List<Category> toRemove = new List<Category>();
                List<Category> allSubs = new List<Category>();
                foreach (Category sub in linesCat.SubCategories)
                    allSubs.Add(sub);
                foreach (Category sub in allSubs)
                {
                    if (sub.Name.StartsWith("HMV_LINEA")) continue;
                    if (IsStandardRevitLineStyle(sub.Name)) continue;

                    GraphicsStyle gs = sub.GetGraphicsStyle(
                        GraphicsStyleType.Projection);
                    if (gs == null) continue;

                    if (!usedStyleIds.Contains(gs.Id))
                        toRemove.Add(sub);
                }

                foreach (Category cat in toRemove)
                {
                    try { doc.Delete(cat.Id); purgedStyles++; }
                    catch { }
                }

                // ===== PURGE UNUSED TEXT TYPES =====
                HashSet<ElementId> usedTextTypeIds = new HashSet<ElementId>(
                    new FilteredElementCollector(doc)
                        .OfClass(typeof(TextNote))
                        .Cast<TextNote>()
                        .Select(tn => tn.GetTypeId()));

                List<TextNoteType> allTextTypes = new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .ToList();

                foreach (TextNoteType tnt in allTextTypes)
                {
                    if (tnt.Name.StartsWith("HMV_General_")) continue;
                    if (usedTextTypeIds.Contains(tnt.Id)) continue;

                    try { doc.Delete(tnt.Id); purgedTextTypes++; }
                    catch { }
                }

                t.Commit();
            }

            TaskDialog.Show("HMV Tools",
                "Texts created: " + textCount + "\n"
                + "Font: " + font + "\n"
                + "Text types: " + typeCache.Count + " HMV types\n\n"
                + "Cleanup:\n"
                + "Unused line styles purged: " + purgedStyles + "\n"
                + "Unused text types purged: " + purgedTextTypes);
            return Result.Succeeded;
        }

        // ===== ACadSharp TEXT EXTRACTION =====

        private void CollectRawText(ACadSharp.Entities.Entity entity,
            ACadSharp.Entities.Insert parentInsert,
            List<double[]> rawCoords, List<DwgTextData> rawTexts)
        {
            string text = null;
            double x = 0, y = 0, height = 0;

            if (entity is AcadText textEnt)
            {
                text = textEnt.Value;
                x = textEnt.InsertPoint.X;
                y = textEnt.InsertPoint.Y;
                height = textEnt.Height;
            }
            else if (entity is AcadMText mtext)
            {
                text = mtext.Value;
                text = StripMTextFormatting(text);
                x = mtext.InsertPoint.X;
                y = mtext.InsertPoint.Y;
                height = mtext.Height;
            }

            if (string.IsNullOrWhiteSpace(text))
                return;

            if (parentInsert != null)
            {
                double ix = parentInsert.InsertPoint.X;
                double iy = parentInsert.InsertPoint.Y;
                double scale = parentInsert.XScale;
                double insRot = parentInsert.Rotation * Math.PI / 180.0;

                double rx = x * scale * Math.Cos(insRot)
                          - y * scale * Math.Sin(insRot) + ix;
                double ry = x * scale * Math.Sin(insRot)
                          + y * scale * Math.Cos(insRot) + iy;

                x = rx;
                y = ry;
                height *= scale;
            }

            rawCoords.Add(new double[] { x, y });
            double heightMm = height * 1000.0;
            rawTexts.Add(new DwgTextData(text,
                new XYZ(x, y, 0), heightMm, 0));
        }

        // ===== DWG BOUNDS (RECURSIVE) =====

        private void CollectAllBounds(ACadSharp.Entities.Entity entity,
            CadDocument cadDoc,
            ref double minX, ref double minY,
            ref double maxX, ref double maxY)
        {
            if (entity is ACadSharp.Entities.Line ln)
            {
                UpdateBounds(ln.StartPoint.X, ln.StartPoint.Y,
                    ref minX, ref minY, ref maxX, ref maxY);
                UpdateBounds(ln.EndPoint.X, ln.EndPoint.Y,
                    ref minX, ref minY, ref maxX, ref maxY);
            }
            else if (entity is ACadSharp.Entities.LwPolyline lw)
            {
                foreach (var v in lw.Vertices)
                    UpdateBounds(v.Location.X, v.Location.Y,
                        ref minX, ref minY, ref maxX, ref maxY);
            }
            else if (entity is ACadSharp.Entities.Polyline2D p2)
            {
                foreach (var v in p2.Vertices)
                    UpdateBounds(v.Location.X, v.Location.Y,
                        ref minX, ref minY, ref maxX, ref maxY);
            }
            else if (entity is ACadSharp.Entities.Circle circ)
            {
                UpdateBounds(circ.Center.X - circ.Radius, circ.Center.Y - circ.Radius,
                    ref minX, ref minY, ref maxX, ref maxY);
                UpdateBounds(circ.Center.X + circ.Radius, circ.Center.Y + circ.Radius,
                    ref minX, ref minY, ref maxX, ref maxY);
            }
            else if (entity is ACadSharp.Entities.Arc arc)
            {
                UpdateBounds(arc.Center.X - arc.Radius, arc.Center.Y - arc.Radius,
                    ref minX, ref minY, ref maxX, ref maxY);
                UpdateBounds(arc.Center.X + arc.Radius, arc.Center.Y + arc.Radius,
                    ref minX, ref minY, ref maxX, ref maxY);
            }
            else if (entity is AcadText txt)
            {
                UpdateBounds(txt.InsertPoint.X, txt.InsertPoint.Y,
                    ref minX, ref minY, ref maxX, ref maxY);
            }
            else if (entity is AcadMText mtxt)
            {
                UpdateBounds(mtxt.InsertPoint.X, mtxt.InsertPoint.Y,
                    ref minX, ref minY, ref maxX, ref maxY);
            }
            else if (entity is ACadSharp.Entities.Insert ins)
            {
                if (ins.Block != null)
                {
                    double ix = ins.InsertPoint.X;
                    double iy = ins.InsertPoint.Y;
                    double scale = ins.XScale;

                    foreach (var be in ins.Block.Entities)
                    {
                        double beMinX = double.MaxValue, beMaxX = double.MinValue;
                        double beMinY = double.MaxValue, beMaxY = double.MinValue;
                        CollectAllBounds(be, cadDoc,
                            ref beMinX, ref beMinY, ref beMaxX, ref beMaxY);

                        if (beMinX != double.MaxValue)
                        {
                            UpdateBounds(beMinX * scale + ix, beMinY * scale + iy,
                                ref minX, ref minY, ref maxX, ref maxY);
                            UpdateBounds(beMaxX * scale + ix, beMaxY * scale + iy,
                                ref minX, ref minY, ref maxX, ref maxY);
                        }
                    }
                }
            }
        }

        private void UpdateBounds(double x, double y,
            ref double minX, ref double minY,
            ref double maxX, ref double maxY)
        {
            if (x < minX) minX = x;
            if (y < minY) minY = y;
            if (x > maxX) maxX = x;
            if (y > maxY) maxY = y;
        }

        // ===== HELPERS =====

        private string StripMTextFormatting(string mtext)
        {
            if (string.IsNullOrEmpty(mtext)) return mtext;

            string result = mtext;
            result = System.Text.RegularExpressions.Regex.Replace(
                result, @"\{\\[^;]+;([^}]*)}", "$1");
            result = System.Text.RegularExpressions.Regex.Replace(
                result, @"\\A\d;", "");
            result = result.Replace("\\P", "\n");
            result = System.Text.RegularExpressions.Regex.Replace(
                result, @"\\[a-zA-Z][^;]*;", "");
            result = result.Replace("{", "").Replace("}", "");

            return result.Trim();
        }

        private string GetDwgFilePath(Document doc, ImportInstance import)
        {
            try
            {
                ElementId typeId = import.GetTypeId();
                Element typeElem = doc.GetElement(typeId);
                if (typeElem != null)
                {
                    ExternalFileReference extRef =
                        typeElem.GetExternalFileReference();
                    if (extRef != null)
                    {
                        ModelPath mp = extRef.GetPath();
                        string path = ModelPathUtils
                            .ConvertModelPathToUserVisiblePath(mp);
                        if (!string.IsNullOrEmpty(path)
                            && System.IO.File.Exists(path))
                            return path;
                    }
                }
            }
            catch { }

            try
            {
                Parameter p = import.get_Parameter(
                    BuiltInParameter.IMPORT_SYMBOL_NAME);
                if (p != null)
                {
                    string name = p.AsString();
                    string[] searchPaths = new string[]
                    {
                        System.IO.Path.GetDirectoryName(
                            doc.PathName ?? ""),
                        Environment.GetFolderPath(
                            Environment.SpecialFolder.Desktop),
                        Environment.GetFolderPath(
                            Environment.SpecialFolder.MyDocuments)
                    };

                    foreach (string dir in searchPaths)
                    {
                        if (string.IsNullOrEmpty(dir)) continue;
                        string full = System.IO.Path.Combine(dir, name);
                        if (!full.EndsWith(".dwg"))
                            full += ".dwg";
                        if (System.IO.File.Exists(full))
                            return full;
                    }
                }
            }
            catch { }

            try
            {
                Parameter p = import.get_Parameter(
                    BuiltInParameter.IMPORT_SYMBOL_NAME);
                string fileName = p != null ? p.AsString() : "DWG file";

                var dialog = new Microsoft.Win32.OpenFileDialog();
                dialog.Title = "Locate: " + fileName;
                dialog.Filter = "DWG files (*.dwg)|*.dwg";
                if (dialog.ShowDialog() == true)
                    return dialog.FileName;
            }
            catch { }

            return null;
        }

        private class DwgTextData
        {
            public string Text;
            public XYZ Position;
            public double HeightMm;
            public double Rotation;

            public DwgTextData(string text, XYZ pos, double hMm, double rot)
            {
                Text = text;
                Position = pos;
                HeightMm = hMm;
                Rotation = rot;
            }
        }

        // ========== LINE HELPERS ==========

        private void GetCurves(Document doc, GeometryElement geoElem,
            List<(Curve curve, GraphicsStyle style)> curveData)
        {
            foreach (GeometryObject geoObj in geoElem)
            {
                GraphicsStyle gs =
                    doc.GetElement(geoObj.GraphicsStyleId) as GraphicsStyle;

                if (geoObj is Curve curve)
                {
                    curveData.Add((curve, gs));
                }
                else if (geoObj is PolyLine polyline)
                {
                    IList<XYZ> pts = polyline.GetCoordinates();
                    for (int i = 0; i < pts.Count - 1; i++)
                    {
                        if (pts[i].DistanceTo(pts[i + 1]) < 0.003) continue;
                        curveData.Add((
                            Autodesk.Revit.DB.Line.CreateBound(
                                pts[i], pts[i + 1]), gs));
                    }
                }
                else if (geoObj is GeometryInstance geoInst)
                {
                    GetCurves(doc, geoInst.GetInstanceGeometry(), curveData);
                }
                else if (geoObj is GeometryElement nestedElem)
                {
                    GetCurves(doc, nestedElem, curveData);
                }
            }
        }

        private string GetHmvLineStyleName(Document doc, GraphicsStyle style)
        {
            if (style == null) return null;
            Category cat = style.GraphicsStyleCategory;
            if (cat == null) return null;

            int? weight = cat.GetLineWeight(GraphicsStyleType.Projection);
            int w = weight.HasValue ? weight.Value : 1;

            ElementId patternId =
                cat.GetLinePatternId(GraphicsStyleType.Projection);
            string patternName = "SOLID";
            if (patternId != null
                && patternId != ElementId.InvalidElementId
                && patternId.IntegerValue > 0)
            {
                LinePatternElement pat =
                    doc.GetElement(patternId) as LinePatternElement;
                if (pat != null) patternName = pat.Name.ToUpper();
            }
            return "HMV_LINEA " + patternName + " " + w;
        }

        private GraphicsStyle CreateLineStyle(Document doc,
            Category linesCat, GraphicsStyle sourceStyle, string name)
        {
            try
            {
                Category newCat = doc.Settings.Categories
                    .NewSubcategory(linesCat, name);
                Category srcCat = sourceStyle.GraphicsStyleCategory;

                int? weight = srcCat.GetLineWeight(
                    GraphicsStyleType.Projection);
                if (weight.HasValue)
                    newCat.SetLineWeight(weight.Value,
                        GraphicsStyleType.Projection);

                ElementId patternId = srcCat.GetLinePatternId(
                    GraphicsStyleType.Projection);
                if (patternId != null
                    && patternId != ElementId.InvalidElementId)
                    newCat.SetLinePatternId(patternId,
                        GraphicsStyleType.Projection);

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

        private TextNoteType CreateTextType(Document doc,
            TextNoteType baseType, string name, double sizeMm, string font)
        {
            try
            {
                TextNoteType nt = baseType.Duplicate(name) as TextNoteType;
                double sizeFeet = sizeMm / 304.8;
                nt.get_Parameter(BuiltInParameter.TEXT_SIZE).Set(sizeFeet);
                nt.get_Parameter(BuiltInParameter.TEXT_FONT).Set(font);
                nt.get_Parameter(BuiltInParameter.TEXT_WIDTH_SCALE).Set(1.0);
                nt.get_Parameter(BuiltInParameter.TEXT_STYLE_BOLD).Set(0);
                nt.get_Parameter(BuiltInParameter.TEXT_STYLE_ITALIC).Set(0);
                nt.get_Parameter(BuiltInParameter.TEXT_BACKGROUND).Set(1);
                return nt;
            }
            catch { return null; }
        }

        private bool IsStandardRevitLineStyle(string name)
        {
            string[] standard = new string[]
            {
                "Thin Lines", "Medium Lines", "Wide Lines", "Lines",
                "<Thin Lines>", "<Medium Lines>", "<Wide Lines>",
                "<Overhead>", "<Beyond>", "<Hidden>", "<Centerline>",
                "<Demolished>", "Axis of Rotation", "Hidden Lines",
                "Demolished", "Overhead", "Beyond", "Centerline",
                "Lineas finas", "Lineas medias", "Lineas anchas",
                "Lineas ocultas", "Eje de rotacion", "Demolido", "Encima"
            };

            foreach (string s in standard)
            {
                if (name.Equals(s, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }
    }

    public enum DwgConvertAction
    {
        ConvertLines,
        StandardizeTexts
    }
}