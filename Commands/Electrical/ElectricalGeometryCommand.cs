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
    // ═══════════════════════════════════════════════════════════════════
    //  COMMAND — ribbon entry point
    // ═══════════════════════════════════════════════════════════════════

    [Transaction(TransactionMode.Manual)]
    public class ElectricalGeometryCommand : IExternalCommand
    {
        private static ElectricalGeometryHandler _handler = null;
        private static ExternalEvent             _exEvent = null;
        private static ElectricalGeometryWindow  _window  = null;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (_window != null && _window.IsLoaded)
            {
                _window.Focus();
                return Result.Succeeded;
            }

            UIApplication uiapp = commandData.Application;
            Document      doc   = uiapp.ActiveUIDocument.Document;

            // Collect FlexPipe parameter names for the dropdown
            var paramNames = new HashSet<string>();
            foreach (FlexPipe fp in new FilteredElementCollector(doc)
                                        .OfClass(typeof(FlexPipe))
                                        .Cast<FlexPipe>())
            {
                foreach (Parameter p in fp.Parameters)
                {
                    try
                    {
                        if (p.Definition != null && !string.IsNullOrWhiteSpace(p.Definition.Name))
                            paramNames.Add(p.Definition.Name);
                    }
                    catch { }
                }
            }
            var sortedParams = new List<string>(paramNames);
            sortedParams.Sort(StringComparer.OrdinalIgnoreCase);

            _handler        = new ElectricalGeometryHandler(doc);
            _exEvent        = ExternalEvent.Create(_handler);

            _window = new ElectricalGeometryWindow(_handler, _exEvent, sortedParams);

            var helper = new System.Windows.Interop.WindowInteropHelper(_window)
            {
                Owner = uiapp.MainWindowHandle
            };

            _window.Show();
            return Result.Succeeded;
        }

        public static void ClearWindow() { _window = null; }
    }

    // ═══════════════════════════════════════════════════════════════════
    //  HANDLER — does the actual work inside a valid Revit API context
    // ═══════════════════════════════════════════════════════════════════

    public class ElectricalGeometryHandler : IExternalEventHandler
    {
        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        private readonly Document _doc;

        // Set by the UI before raising the event
        public ElectricalGeometryWindow UI           { get; set; }
        public string                   TargetParam  { get; set; }
        public bool                     IncludeDxf   { get; set; }
        public string                   DxfFolder    { get; set; }
        public string                   OutputPath   { get; set; }

        public ElectricalGeometryHandler(Document doc)
        {
            _doc = doc;
        }

        public string GetName() => "ElectricalGeometryHandler";

        public void Execute(UIApplication app)
        {
            try
            {
                UI?.SetStatus("Collecting FlexPipe data...");

                // ── 1. Optionally scan DXF files ─────────────────────
                var dxfMap      = new Dictionary<string, string>();
                var dxfGeometry = new Dictionary<string, (double vano, double desnivel)>();

                if (IncludeDxf && !string.IsNullOrWhiteSpace(DxfFolder))
                {
                    UI?.SetStatus("Scanning DXF files...");
                    foreach (string filePath in Directory.GetFiles(DxfFolder, "*.dxf", SearchOption.AllDirectories))
                    {
                        string fileName   = Path.GetFileNameWithoutExtension(filePath);
                        int    uIdx       = fileName.IndexOf('_');
                        string matchKey   = uIdx > 0 ? fileName.Substring(0, uIdx) : fileName;
                        if (!dxfMap.ContainsKey(matchKey))
                            dxfMap.Add(matchKey, filePath);
                    }

                    foreach (var kvp in dxfMap)
                    {
                        ExtractDxfGeometry(kvp.Value, out double v, out double d);
                        dxfGeometry[kvp.Key] = (v, d);
                    }
                    UI?.SetStatus($"  {dxfMap.Count} DXF file(s) found.");
                }

                // ── 2. Collect FlexPipe rows ─────────────────────────
                double ftToMm = UnitUtils.ConvertFromInternalUnits(1.0, UnitTypeId.Millimeters);

                var dataRows    = new List<(string paramValue, string equipoInicial, string equipoFinal,
                                            double? dxfVano, double? dxfDesnivel,
                                            double rvtVano, double rvtDesnivel)>();
                var matchedKeys = new HashSet<string>();
                int skipped     = 0;

                foreach (FlexPipe flex in new FilteredElementCollector(_doc)
                                              .OfClass(typeof(FlexPipe))
                                              .Cast<FlexPipe>())
                {
                    Parameter param = flex.LookupParameter(TargetParam);
                    if (param == null || !param.HasValue) { skipped++; continue; }

                    string paramValue = param.AsString();
                    if (string.IsNullOrWhiteSpace(paramValue)) { skipped++; continue; }

                    IList<XYZ> pts = flex.Points;
                    if (pts == null || pts.Count < 2) { skipped++; continue; }

                    XYZ pA = pts.First();
                    XYZ pB = pts.Last();

                    double rvtDx       = (pB.X - pA.X) * ftToMm;
                    double rvtDy       = (pB.Y - pA.Y) * ftToMm;
                    double rvtDz       = (pB.Z - pA.Z) * ftToMm;
                    double rvtVano     = Math.Sqrt(rvtDx * rvtDx + rvtDy * rvtDy);
                    double rvtDesnivel = Math.Abs(rvtDz);

                    Parameter pInicial    = flex.LookupParameter("HMV_CFI_EQUIPO INICIAL");
                    Parameter pFinal      = flex.LookupParameter("HMV_CFI_EQUIPO FINAL");
                    string    equipInicial = pInicial != null && pInicial.HasValue ? pInicial.AsString() ?? "" : "";
                    string    equipFinal   = pFinal   != null && pFinal.HasValue   ? pFinal.AsString()   ?? "" : "";

                    double? dxfVano     = null;
                    double? dxfDesnivel = null;

                    if (IncludeDxf && dxfGeometry.ContainsKey(paramValue))
                    {
                        dxfVano     = dxfGeometry[paramValue].vano;
                        dxfDesnivel = dxfGeometry[paramValue].desnivel;
                        matchedKeys.Add(paramValue);
                    }

                    dataRows.Add((paramValue, equipInicial, equipFinal, dxfVano, dxfDesnivel, rvtVano, rvtDesnivel));
                }

                UI?.SetStatus($"  {dataRows.Count} FlexPipe(s) matched, {skipped} skipped (no parameter value).");

                if (dataRows.Count == 0)
                {
                    UI?.SetStatus($"Warning: No FlexPipes found with parameter \"{TargetParam}\". Check the parameter name.");
                    return;
                }

                // ── 3. Sort ──────────────────────────────────────────
                dataRows.Sort((a, b) => string.Compare(a.paramValue, b.paramValue, StringComparison.OrdinalIgnoreCase));

                // ── 4. Build CSV ─────────────────────────────────────
                var csvLines = new List<string>();

                if (IncludeDxf)
                    csvLines.Add("PARAMETER NAME;HMV_CFI_EQUIPO INICIAL;HMV_CFI_EQUIPO FINAL;dxf vano (mm);dxf desnivel (mm);rvt vano (mm);rvt desnivel (mm);vano diff (mm);desnivel diff (mm)");
                else
                    csvLines.Add("PARAMETER NAME;HMV_CFI_EQUIPO INICIAL;HMV_CFI_EQUIPO FINAL;rvt vano (mm);rvt desnivel (mm)");

                foreach (var row in dataRows)
                {
                    if (IncludeDxf)
                    {
                        string dxfVanoStr     = row.dxfVano.HasValue     ? row.dxfVano.Value.ToString("F4", Inv)     : "";
                        string dxfDesnivelStr = row.dxfDesnivel.HasValue ? row.dxfDesnivel.Value.ToString("F4", Inv) : "";
                        string vanoDiff       = row.dxfVano.HasValue     ? Math.Abs(row.dxfVano.Value     - row.rvtVano).ToString("F4", Inv)     : "";
                        string desnivelDiff   = row.dxfDesnivel.HasValue ? Math.Abs(row.dxfDesnivel.Value - row.rvtDesnivel).ToString("F4", Inv) : "";

                        csvLines.Add(string.Join(";",
                            row.paramValue, row.equipoInicial, row.equipoFinal,
                            dxfVanoStr, dxfDesnivelStr,
                            row.rvtVano.ToString("F4", Inv), row.rvtDesnivel.ToString("F4", Inv),
                            vanoDiff, desnivelDiff));
                    }
                    else
                    {
                        csvLines.Add(string.Join(";",
                            row.paramValue, row.equipoInicial, row.equipoFinal,
                            row.rvtVano.ToString("F4", Inv), row.rvtDesnivel.ToString("F4", Inv)));
                    }
                }

                // ── 5. Append unmatched DXFs ─────────────────────────
                if (IncludeDxf)
                {
                    foreach (var kvp in dxfMap)
                    {
                        if (matchedKeys.Contains(kvp.Key)) continue;
                        ExtractDxfGeometry(kvp.Value, out double rv, out double rd);
                        csvLines.Add(string.Join(";",
                            "dxf_notfound_" + Path.GetFileNameWithoutExtension(kvp.Value),
                            "", "", rv.ToString("F4", Inv), rd.ToString("F4", Inv),
                            "", "", "", ""));
                    }
                }

                // ── 6. Save CSV directly — path was chosen on the UI thread ─
                File.WriteAllLines(OutputPath, csvLines);
                UI?.SetStatus("✔ Report saved to: " + OutputPath);
                UI?.SetStatus("  Total rows: " + dataRows.Count);
            }
            catch (Exception ex)
            {
                UI?.SetStatus("Error: " + ex.Message);
            }
        }

        private static void ExtractDxfGeometry(string filepath, out double vano, out double desnivel)
        {
            vano = 0;
            desnivel = 0;
            try
            {
                DxfFile dxf      = DxfFile.Load(filepath);
                var     entities = dxf.Entities.Where(e => e.Layer == "0").ToList();

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
                        var s = lwPoly.Vertices.First();
                        var e2 = lwPoly.Vertices.Last();
                        ptStart = new DxfPoint(s.X, s.Y, 0);
                        ptEnd   = new DxfPoint(e2.X, e2.Y, 0);
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
            catch { }
        }
    }
}
