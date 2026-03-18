using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;

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
            {
                return ConvertLines(doc, view, selectedDwgs);
            }
            else if (win.Action == DwgConvertAction.StandardizeAll)
            {
                return StandardizeAll(doc, view, font);
            }

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
                + "Next: Partial Explode the DWG(s), then run again "
                + "and click 'Standardize Texts'.");
            return Result.Succeeded;
        }

        // ========== STANDARDIZE ALL (lines + texts after Partial Explode) ==========
        private static readonly double[] AllowedSizes =
            { 1.5, 2.0, 2.5, 3.0, 3.5 };

        private Result StandardizeAll(Document doc, View view, string font)
        {
            int lineCurveCount = 0;
            int textCount = 0;
            int purgedStyles = 0;
            int purgedTextTypes = 0;
            var lineStyleCache = new Dictionary<string, GraphicsStyle>();
            var typeCache = new Dictionary<string, TextNoteType>();

            Category linesCat = doc.Settings.Categories
                .get_Item(BuiltInCategory.OST_Lines);

            // Pre-load existing HMV line styles
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

            // Pre-load existing HMV text types
            foreach (TextNoteType tnt in new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .ToList())
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
                TaskDialog.Show("HMV Tools", "No text types found in document.");
                return Result.Cancelled;
            }

            using (Transaction t = new Transaction(doc,
                "Standardize Lines & Texts"))
            {
                t.Start();

                // ===== 1. STANDARDIZE DETAIL LINES (from Partial Explode) =====
                // Re-styles DetailCurves that have DWG-imported line styles

                List<DetailCurve> curves = new FilteredElementCollector(doc, view.Id)
                    .OfClass(typeof(CurveElement))
                    .OfType<DetailCurve>()
                    .ToList();

                foreach (DetailCurve dc in curves)
                {
                    try
                    {
                        GraphicsStyle gs = dc.LineStyle as GraphicsStyle;
                        if (gs == null) continue;
                        if (gs.Name.StartsWith("HMV_LINEA")) continue;
                        if (IsStandardRevitLineStyle(gs.Name)) continue;

                        string styleName = GetHmvLineStyleName(doc, gs);
                        if (styleName == null) continue;

                        if (!lineStyleCache.TryGetValue(styleName,
                            out GraphicsStyle hmvStyle))
                        {
                            hmvStyle = CreateLineStyle(
                                doc, linesCat, gs, styleName);
                            if (hmvStyle != null)
                                lineStyleCache[styleName] = hmvStyle;
                        }

                        if (hmvStyle != null)
                        {
                            dc.LineStyle = hmvStyle;
                            lineCurveCount++;
                        }
                    }
                    catch { }
                }

                // ===== 2. STANDARDIZE TEXTS =====

                List<TextNote> texts = new FilteredElementCollector(doc, view.Id)
                    .OfClass(typeof(TextNote))
                    .Cast<TextNote>()
                    .ToList();

                foreach (TextNote tn in texts)
                {
                    try
                    {
                        TextNoteType currentType =
                            doc.GetElement(tn.GetTypeId()) as TextNoteType;
                        if (currentType == null) continue;
                        if (currentType.Name.StartsWith("HMV_General_")) continue;

                        double sizeFeet = currentType
                            .get_Parameter(BuiltInParameter.TEXT_SIZE)
                            .AsDouble();
                        double sizeMm = sizeFeet * 304.8;
                        double snapped = SnapToAllowed(sizeMm);

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
                            tn.ChangeTypeId(hmvType.Id);

                        textCount++;
                    }
                    catch { }
                }

                // ===== 3. PURGE UNUSED IMPORTED LINE STYLES =====
                // Collect used style IDs first, then delete — never modify
                // a live FilteredElementCollector during iteration.

                HashSet<ElementId> usedStyleIds = new HashSet<ElementId>();
                foreach (CurveElement ce in new FilteredElementCollector(doc)
                    .OfClass(typeof(CurveElement))
                    .Cast<CurveElement>()
                    .ToList())
                {
                    try
                    {
                        GraphicsStyle gs = ce.LineStyle as GraphicsStyle;
                        if (gs != null)
                            usedStyleIds.Add(gs.Id);
                    }
                    catch { }
                }

                List<ElementId> lineStyleIdsToDelete = new List<ElementId>();
                foreach (Category sub in linesCat.SubCategories)
                {
                    if (sub.Name.StartsWith("HMV_LINEA")) continue;
                    if (IsStandardRevitLineStyle(sub.Name)) continue;

                    GraphicsStyle gs = sub.GetGraphicsStyle(
                        GraphicsStyleType.Projection);
                    if (gs == null) continue;

                    if (!usedStyleIds.Contains(gs.Id))
                        lineStyleIdsToDelete.Add(sub.Id);
                }

                foreach (ElementId id in lineStyleIdsToDelete)
                {
                    try { doc.Delete(id); purgedStyles++; }
                    catch { }
                }

                // ===== 4. PURGE UNUSED TEXT TYPES =====
                // Materialize the collector to a List BEFORE calling doc.Delete
                // inside the loop — mutating the element table while iterating
                // a live collector is what caused the "iterator cannot proceed"
                // crash.

                HashSet<ElementId> usedTextTypeIds = new HashSet<ElementId>(
                    new FilteredElementCollector(doc)
                        .OfClass(typeof(TextNote))
                        .Cast<TextNote>()
                        .Select(note => note.GetTypeId()));

                List<ElementId> textTypeIdsToDelete = new List<ElementId>();
                foreach (TextNoteType tnt in new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .ToList())
                {
                    if (tnt.Name.StartsWith("HMV_General_")) continue;
                    if (usedTextTypeIds.Contains(tnt.Id)) continue;
                    textTypeIdsToDelete.Add(tnt.Id);
                }

                foreach (ElementId id in textTypeIdsToDelete)
                {
                    try { doc.Delete(id); purgedTextTypes++; }
                    catch { }
                }

                t.Commit();
            }

            TaskDialog.Show("HMV Tools",
                "Lines standardized: " + lineCurveCount + "\n"
                + "Texts standardized: " + textCount + "\n"
                + "Font: " + font + "\n\n"
                + "Cleanup:\n"
                + "Unused line styles purged: " + purgedStyles + "\n"
                + "Unused text types purged: " + purgedTextTypes);
            return Result.Succeeded;
        }

        // ========== HELPERS ==========
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
                            Line.CreateBound(pts[i], pts[i + 1]), gs));
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

                newCat.LineColor = new Color(0, 0, 0);
                return newCat.GetGraphicsStyle(GraphicsStyleType.Projection);
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
                "Líneas finas", "Líneas medias", "Líneas anchas",
                "Líneas ocultas", "Eje de rotación", "Demolido", "Encima"
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
        StandardizeAll
    }
}