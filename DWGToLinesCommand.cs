using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class DwgToLinesCommand : IExternalCommand
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
                TaskDialog.Show("DWG", "No DWG imports found.");
                return Result.Cancelled;
            }

            List<string> names = new List<string>();
            foreach (var d in dwgs)
            {
                string name = d.Category != null ? d.Category.Name : "Unknown";
                names.Add(name + " [ID: " + d.Id.IntegerValue + "]");
            }

            DwgPickerWindow picker = new DwgPickerWindow(names);
            if (picker.ShowDialog() != true)
                return Result.Cancelled;

            List<ImportInstance> selectedDwgs = new List<ImportInstance>();
            foreach (int idx in picker.SelectedIndices)
                selectedDwgs.Add(dwgs[idx]);

            int count = 0;
            Category linesCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
            var lineStyleCache = new Dictionary<string, GraphicsStyle>();

            using (Transaction t = new Transaction(doc, "DWG to Detail Lines"))
            {
                t.Start();

                // Pre-load existing HMV line styles
                foreach (Category subCat in linesCat.SubCategories)
                {
                    if (subCat.Name.StartsWith("HMV_LINEA"))
                    {
                        GraphicsStyle gs = subCat.GetGraphicsStyle(GraphicsStyleType.Projection);
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
                            if (item.curve.Length < 0.003)
                                continue;

                            DetailCurve dc = doc.Create.NewDetailCurve(view, item.curve);

                            string styleName = GetHmvLineStyleName(doc, item.style);
                            if (styleName != null)
                            {
                                if (!lineStyleCache.TryGetValue(styleName, out GraphicsStyle hmvStyle))
                                {
                                    hmvStyle = CreateLineStyle(doc, linesCat, item.style, styleName);
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

            TaskDialog.Show("DWG",
                count + " detail lines created from "
                + selectedDwgs.Count + " DWG(s).\n"
                + lineStyleCache.Count + " HMV line styles used.");
            return Result.Succeeded;
        }

        private void GetCurves(Document doc, GeometryElement geoElem,
            List<(Curve curve, GraphicsStyle style)> curveData)
        {
            foreach (GeometryObject geoObj in geoElem)
            {
                GraphicsStyle gs = doc.GetElement(geoObj.GraphicsStyleId) as GraphicsStyle;

                if (geoObj is Curve curve)
                {
                    curveData.Add((curve, gs));
                }
                else if (geoObj is PolyLine polyline)
                {
                    IList<XYZ> pts = polyline.GetCoordinates();
                    for (int i = 0; i < pts.Count - 1; i++)
                    {
                        if (pts[i].DistanceTo(pts[i + 1]) < 0.003)
                            continue;
                        curveData.Add((Line.CreateBound(pts[i], pts[i + 1]), gs));
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
            if (style == null)
                return null;

            Category cat = style.GraphicsStyleCategory;
            if (cat == null)
                return null;

            int? weight = cat.GetLineWeight(GraphicsStyleType.Projection);
            int w = weight.HasValue ? weight.Value : 1;

            ElementId patternId = cat.GetLinePatternId(GraphicsStyleType.Projection);
            string patternName = "SOLID";
            if (patternId != null
                && patternId != ElementId.InvalidElementId
                && patternId.IntegerValue > 0)
            {
                LinePatternElement patElem = doc.GetElement(patternId) as LinePatternElement;
                if (patElem != null)
                    patternName = patElem.Name.ToUpper();
            }

            return "HMV_LINEA " + patternName + " " + w;
        }

        private GraphicsStyle CreateLineStyle(Document doc, Category linesCat,
            GraphicsStyle sourceStyle, string name)
        {
            try
            {
                Category newCat = doc.Settings.Categories.NewSubcategory(linesCat, name);
                Category srcCat = sourceStyle.GraphicsStyleCategory;

                int? weight = srcCat.GetLineWeight(GraphicsStyleType.Projection);
                if (weight.HasValue)
                    newCat.SetLineWeight(weight.Value, GraphicsStyleType.Projection);

                ElementId patternId = srcCat.GetLinePatternId(GraphicsStyleType.Projection);
                if (patternId != null && patternId != ElementId.InvalidElementId)
                    newCat.SetLinePatternId(patternId, GraphicsStyleType.Projection);

                newCat.LineColor = new Color(0, 0, 0);

                return newCat.GetGraphicsStyle(GraphicsStyleType.Projection);
            }
            catch { return null; }
        }
    }
}