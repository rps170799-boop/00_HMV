using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using IxMilia.Dxf;
using IxMilia.Dxf.Entities;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class ElectricalGeometryCommand : IExternalCommand
    {
        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            // 1. Show UI
            ElectricalGeometryWindow window = new ElectricalGeometryWindow();
            var helper = new System.Windows.Interop.WindowInteropHelper(window)
            {
                Owner = uiapp.MainWindowHandle
            };

            window.ShowDialog();

            // 2. Validate run
            if (!window.Proceed)
                return Result.Cancelled;

            string targetParam = window.ParameterName;
            string dxfFolder   = window.SelectedFolder;

            try
            {
                // 3. Scan DXF Files
                var dxfFiles = Directory.GetFiles(dxfFolder, "*.dxf", SearchOption.AllDirectories);
                Dictionary<string, string> dxfMap = new Dictionary<string, string>();

                foreach (string filePath in dxfFiles)
                {
                    string fileName   = Path.GetFileNameWithoutExtension(filePath);
                    int underscoreIdx = fileName.IndexOf('_');
                    string matchKey   = underscoreIdx > 0 ? fileName.Substring(0, underscoreIdx) : fileName;

                    if (!dxfMap.ContainsKey(matchKey))
                        dxfMap.Add(matchKey, filePath);
                }

                // Pre-extract DXF geometry — coords are already in mm, NO division
                Dictionary<string, (double vano, double desnivel)> dxfGeometry =
                    new Dictionary<string, (double, double)>();

                foreach (var kvp in dxfMap)
                {
                    ExtractDxfGeometry(kvp.Value, out double vano, out double desnivel);
                    dxfGeometry[kvp.Key] = (vano, desnivel); // raw mm, no scaling
                }

                // 4. Get all Flex Pipes in model
                var flexPipes = new FilteredElementCollector(doc)
                    .OfClass(typeof(FlexPipe))
                    .Cast<FlexPipe>()
                    .ToList();

                // Revit internal unit = feet → mm: 1 ft = 304.8 mm, no extra factor
                double ftToMm = UnitUtils.ConvertFromInternalUnits(1.0, UnitTypeId.Millimeters);

                var dataRows = new List<(string paramValue, double? dxfVano, double? dxfDesnivel,
                                         double rvtVano, double rvtDesnivel)>();

                HashSet<string> matchedKeys = new HashSet<string>();

                // 5. Collect rows — DXF value repeated on EVERY row of the same group
                foreach (FlexPipe flex in flexPipes)
                {
                    Parameter param = flex.LookupParameter(targetParam);
                    if (param == null || !param.HasValue) continue;

                    string paramValue = param.AsString();
                    if (string.IsNullOrWhiteSpace(paramValue)) continue;

                    IList<XYZ> pts = flex.Points;
                    if (pts == null || pts.Count < 2) continue;

                    XYZ pA = pts.First();
                    XYZ pB = pts.Last();

                    double rvtDx = (pB.X - pA.X) * ftToMm;
                    double rvtDy = (pB.Y - pA.Y) * ftToMm;
                    double rvtDz = (pB.Z - pA.Z) * ftToMm;

                    double rvtVano     = Math.Sqrt(rvtDx * rvtDx + rvtDy * rvtDy);
                    double rvtDesnivel = Math.Abs(rvtDz);

                    double? dxfVano     = null;
                    double? dxfDesnivel = null;

                    if (dxfGeometry.ContainsKey(paramValue))
                    {
                        dxfVano     = dxfGeometry[paramValue].vano;
                        dxfDesnivel = dxfGeometry[paramValue].desnivel;
                        matchedKeys.Add(paramValue);
                    }

                    dataRows.Add((paramValue, dxfVano, dxfDesnivel, rvtVano, rvtDesnivel));
                }

                // 6. Sort alphabetically by PARAMETER NAME
                dataRows.Sort((a, b) => string.Compare(a.paramValue, b.paramValue, StringComparison.OrdinalIgnoreCase));

                // 7. Build CSV — F4 to preserve small decimal differences
                List<string> csvLines = new List<string>();
                csvLines.Add("PARAMETER NAME;dxf vano (mm);dxf desnivel (mm);rvt vano (mm);rvt desnivel (mm);vano diff (mm);desnivel diff (mm)");

                foreach (var row in dataRows)
                {
                    string dxfVanoStr     = row.dxfVano.HasValue
                        ? row.dxfVano.Value.ToString("F4", Inv)     : "";
                    string dxfDesnivelStr = row.dxfDesnivel.HasValue
                        ? row.dxfDesnivel.Value.ToString("F4", Inv) : "";

                    string vanoDiffStr     = row.dxfVano.HasValue
                        ? Math.Abs(row.dxfVano.Value     - row.rvtVano).ToString("F4", Inv)     : "";
                    string desnivelDiffStr = row.dxfDesnivel.HasValue
                        ? Math.Abs(row.dxfDesnivel.Value - row.rvtDesnivel).ToString("F4", Inv) : "";

                    csvLines.Add(string.Join(";",
                        row.paramValue,
                        dxfVanoStr,
                        dxfDesnivelStr,
                        row.rvtVano.ToString("F4", Inv),
                        row.rvtDesnivel.ToString("F4", Inv),
                        vanoDiffStr,
                        desnivelDiffStr));
                }

                // 8. Append Unmatched DXFs (no FlexPipe found for them)
                foreach (var kvp in dxfMap)
                {
                    if (matchedKeys.Contains(kvp.Key)) continue;

                    ExtractDxfGeometry(kvp.Value, out double rawVano, out double rawDesnivel);

                    string fullFileName = Path.GetFileNameWithoutExtension(kvp.Value);
                    csvLines.Add(string.Join(";",
                        "dxf_notfound_" + fullFileName,
                        rawVano.ToString("F4", Inv),
                        rawDesnivel.ToString("F4", Inv),
                        "", "", "", ""));
                }

                // 9. Save CSV
                using (var saveDialog = new System.Windows.Forms.SaveFileDialog())
                {
                    saveDialog.Filter   = "CSV Files (*.csv)|*.csv";
                    saveDialog.Title    = "Save Electrical Geometry Report";
                    saveDialog.FileName = "ElectricalGeometryReport.csv";

                    if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        File.WriteAllLines(saveDialog.FileName, csvLines);
                        TaskDialog.Show("Success", "Report exported successfully to:\n" + saveDialog.FileName);
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", "An error occurred: " + ex.Message);
                return Result.Failed;
            }
        }

        private void ExtractDxfGeometry(string filepath, out double vano, out double desnivel)
        {
            vano = 0;
            desnivel = 0;

            try
            {
                DxfFile dxf  = DxfFile.Load(filepath);
                var entities = dxf.Entities.Where(e => e.Layer == "0").ToList();

                DxfPoint? ptStart = null;
                DxfPoint? ptEnd   = null;

                foreach (var ent in entities)
                {
                    if (ent is DxfPolyline poly && poly.Vertices.Count > 1)
                    {
                        ptStart = poly.Vertices.First().Location;
                        ptEnd   = poly.Vertices.Last().Location;
                        break;
                    }
                    else if (ent is DxfLwPolyline lwPoly && lwPoly.Vertices.Count > 1)
                    {
                        var start = lwPoly.Vertices.First();
                        var end   = lwPoly.Vertices.Last();
                        ptStart   = new DxfPoint(start.X, start.Y, 0);
                        ptEnd     = new DxfPoint(end.X,   end.Y,   0);
                        break;
                    }
                    else if (ent is DxfLine line)
                    {
                        ptStart = line.P1;
                        ptEnd   = line.P2;
                        break;
                    }
                }

                if (ptStart.HasValue && ptEnd.HasValue)
                {
                    vano     = Math.Abs(ptEnd.Value.X - ptStart.Value.X);
                    desnivel = Math.Abs(ptEnd.Value.Y - ptStart.Value.Y);
                }
            }
            catch
            {
                // Fail silently, leave 0
            }
        }
    }
}