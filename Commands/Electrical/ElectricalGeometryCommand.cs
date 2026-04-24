using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using IxMilia.Dxf;
using IxMilia.Dxf.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class ElectricalGeometryCommand : IExternalCommand
    {
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
            string dxfFolder = window.SelectedFolder;

            try
            {
                // 3. Scan DXF Files
                var dxfFiles = Directory.GetFiles(dxfFolder, "*.dxf", SearchOption.AllDirectories);
                Dictionary<string, string> dxfMap = new Dictionary<string, string>(); // Key=StrippedName, Value=FilePath

                foreach (string filePath in dxfFiles)
                {
                    string fileName = Path.GetFileNameWithoutExtension(filePath);
                    int underscoreIdx = fileName.IndexOf('_');
                    string matchKey = underscoreIdx > 0 ? fileName.Substring(0, underscoreIdx) : fileName;

                    if (!dxfMap.ContainsKey(matchKey))
                        dxfMap.Add(matchKey, filePath);
                }
                Dictionary<string, (double vano, double desnivel)> dxfGeometry = new Dictionary<string, (double, double)>();
                foreach (var kvp in dxfMap)
                {
                    ExtractDxfGeometry(kvp.Value, out double vano, out double desnivel);
                    dxfGeometry[kvp.Key] = (vano, desnivel);
                }




                // 4. Get all Flex Pipes in model
                var flexPipes = new FilteredElementCollector(doc)
                    .OfClass(typeof(FlexPipe))
                    .Cast<FlexPipe>()
                    .ToList();

                List<string> csvLines = new List<string>();
                csvLines.Add("PARAMETER NAME ; dxf vano; dxf desnivel; rvt vano; rvt desnivel");

                double ftToMm = UnitUtils.ConvertFromInternalUnits(1.0, UnitTypeId.Millimeters);

                HashSet<string> matchedKeys = new HashSet<string>();

                // 5. Compare and extract
                foreach (FlexPipe flex in flexPipes)
                {
                    Parameter param = flex.LookupParameter(targetParam);
                    if (param == null || !param.HasValue) continue;

                    string paramValue = param.AsString();
                    if (string.IsNullOrWhiteSpace(paramValue)) continue; // Rule: if empty, shouldn't enter

                    // Calculate Revit Vano/Desnivel
                    IList<XYZ> pts = flex.Points;
                    if (pts == null || pts.Count < 2) continue;

                    XYZ pA = pts.First();
                    XYZ pB = pts.Last();

                    double rvtDx = (pB.X - pA.X) * ftToMm;
                    double rvtDy = (pB.Y - pA.Y) * ftToMm;
                    double rvtDz = (pB.Z - pA.Z) * ftToMm;

                    double rvtVano = Math.Sqrt(rvtDx * rvtDx + rvtDy * rvtDy);
                    double rvtDesnivel = Math.Abs(rvtDz);

                    string dxfVanoStr = "";
                    string dxfDesnivelStr = "";

                    // Find matching DXF
                    if (dxfGeometry.ContainsKey(paramValue))
                    {
                        dxfVanoStr = dxfGeometry[paramValue].vano.ToString("F2");
                        dxfDesnivelStr = dxfGeometry[paramValue].desnivel.ToString("F2");    
                        matchedKeys.Add(paramValue);
                        // Remove from map to track "Not Found" ones later
                        dxfGeometry.Remove(paramValue);
                    }

                    // Output row
                    csvLines.Add($"{paramValue} ; {dxfVanoStr} ; {dxfDesnivelStr} ; {rvtVano:F2} ; {rvtDesnivel:F2}");
                }

                // 6. Append Unmatched DXFs at the end
                foreach (var kvp in dxfMap)
                {
                    string matchKey = kvp.Key;
                    string dxfPath = kvp.Value;
                    string fullFileName = Path.GetFileNameWithoutExtension(dxfPath);

                    ExtractDxfGeometry(dxfPath, out double dxfVano, out double dxfDesnivel);

                    csvLines.Add($"dxf_notfound_{fullFileName} ; {dxfVano:F2} ; {dxfDesnivel:F2} ; ; ");
                }

                // 7. Save CSV File
                using (var saveDialog = new System.Windows.Forms.SaveFileDialog())
                {
                    saveDialog.Filter = "CSV Files (*.csv)|*.csv";
                    saveDialog.Title = "Save Electrical Geometry Report";
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

        // ── DXF Geometry Parsing (Reusing your logic) ──
        // ── DXF Geometry Parsing (Reusing your logic) ──
        private void ExtractDxfGeometry(string filepath, out double vano, out double desnivel)
        {
            vano = 0;
            desnivel = 0;

            try
            {
                DxfFile dxf = DxfFile.Load(filepath);
                var entities = dxf.Entities.Where(e => e.Layer == "0").ToList();

                // FIX: Make DxfPoint nullable by adding '?'
                DxfPoint? ptStart = null;
                DxfPoint? ptEnd = null;

                foreach (var ent in entities)
                {
                    if (ent is DxfPolyline poly && poly.Vertices.Count > 1)
                    {
                        ptStart = poly.Vertices.First().Location;
                        ptEnd = poly.Vertices.Last().Location;
                        break;
                    }
                    else if (ent is DxfLwPolyline lwPoly && lwPoly.Vertices.Count > 1)
                    {
                        var start = lwPoly.Vertices.First();
                        var end = lwPoly.Vertices.Last();
                        ptStart = new DxfPoint(start.X, start.Y, 0);
                        ptEnd = new DxfPoint(end.X, end.Y, 0);
                        break;
                    }
                    else if (ent is DxfLine line)
                    {
                        ptStart = line.P1;
                        ptEnd = line.P2;
                        break;
                    }
                }

                // FIX: Check .HasValue and use .Value
                if (ptStart.HasValue && ptEnd.HasValue)
                {
                    // DX is horizontal distance, DY is vertical distance in CAD mapping
                    vano = Math.Abs(ptEnd.Value.X - ptStart.Value.X);
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