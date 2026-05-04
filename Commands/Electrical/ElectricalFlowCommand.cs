using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

namespace HMVTools
{
    // ═══════════════════════════════════════════════════════════════════
    //  CONFIGURATION — set by the Config window, consumed by the handler
    // ═══════════════════════════════════════════════════════════════════

    public class ElectricalFlowConfig
    {
        public string SourceEquipmentParam { get; set; }
        public string DestPointParam       { get; set; }
    }

    // ═══════════════════════════════════════════════════════════════════
    //  COMMAND
    // ═══════════════════════════════════════════════════════════════════

    [Transaction(TransactionMode.Manual)]
    public class ElectricalFlowCommand : IExternalCommand
    {
        private static ElectricalFlowPickHandler   _pickHandler   = null;
        private static ExternalEvent               _pickEvent     = null;
        private static ElectricalFlowExportHandler _exportHandler = null;
        private static ExternalEvent               _exportEvent   = null;
        private static ElectricalFlowWindow        _window        = null;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (_window != null && _window.IsLoaded)
            {
                _window.Focus();
                return Result.Succeeded;
            }

            _pickHandler   = new ElectricalFlowPickHandler();
            _pickEvent     = ExternalEvent.Create(_pickHandler);
            _exportHandler = new ElectricalFlowExportHandler();
            _exportEvent   = ExternalEvent.Create(_exportHandler);

            _window = new ElectricalFlowWindow(
                commandData.Application,
                _pickHandler,   _pickEvent,
                _exportHandler, _exportEvent);

            var helper = new System.Windows.Interop.WindowInteropHelper(_window);
            helper.Owner = commandData.Application.MainWindowHandle;

            _window.Show();
            return Result.Succeeded;
        }

        public static void ClearWindow() { _window = null; }
    }

    // ═══════════════════════════════════════════════════════════════════
    //  PICK HANDLER  — hides the window, picks adaptive components,
    //                  returns ids to the UI via OnPointsPicked()
    // ═══════════════════════════════════════════════════════════════════

    public class ElectricalFlowPickHandler : IExternalEventHandler
    {
        public ElectricalFlowWindow UI { get; set; }

        public string GetName() => "ElectricalFlowPickHandler";

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;

            try
            {
                IList<Reference> refs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    new AdaptiveComponentFilter(),
                    "Select Adaptive Components — click Finish when done");

                if (refs == null || refs.Count == 0)
                {
                    UI?.SetStatus("No elements selected.");
                    UI?.RestoreWindow();
                    return;
                }

                UI?.OnPointsPicked(refs.Select(r => r.ElementId).ToList());
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                UI?.SetStatus("Selection cancelled.");
                UI?.RestoreWindow();
            }
            catch (Exception ex)
            {
                UI?.SetStatus("Pick error: " + ex.Message);
                UI?.RestoreWindow();
            }
        }
    }

    // ═══════════════════════════════════════════════════════════════════
    //  EXPORT HANDLER  — finds closest equipment per point, injects param
    // ═══════════════════════════════════════════════════════════════════

    public class ElectricalFlowExportHandler : IExternalEventHandler
    {
        public ElectricalFlowWindow  UI       { get; set; }
        public List<ElementId>       PointIds { get; set; }
        public ElectricalFlowConfig  Config   { get; set; }

        private static readonly System.Globalization.CultureInfo Inv =
            System.Globalization.CultureInfo.InvariantCulture;

        public string GetName() => "ElectricalFlowExportHandler";

        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;

            if (PointIds == null || PointIds.Count == 0 || Config == null)
            {
                UI?.SetStatus("DBG Error: No points or config. PointIds=" +
                    (PointIds == null ? "null" : PointIds.Count.ToString()) +
                    " Config=" + (Config == null ? "null" : "ok"));
                return;
            }

            if (string.IsNullOrWhiteSpace(Config.SourceEquipmentParam) ||
                string.IsNullOrWhiteSpace(Config.DestPointParam))
            {
                UI?.SetStatus("DBG Error: Params not configured. Source='" +
                    Config.SourceEquipmentParam + "' Dest='" + Config.DestPointParam + "'");
                return;
            }

            try
            {
                UI?.SetStatus($"DBG Step 1: Collecting Electrical Equipment... ({PointIds.Count} points to process)");

                // ── 1. Collect all Electrical Equipment once ──────────────────
                var allEquipment = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .ToList();

                UI?.SetStatus($"DBG Step 1: Found {allEquipment.Count} equipment total.");

                if (allEquipment.Count == 0)
                {
                    UI?.SetStatus("DBG Error: No Electrical Equipment found in the model.");
                    return;
                }

                // ── 1b. Filter to only equipment that have a non-empty source param ──
                var candidateEquipment = allEquipment
                    .Where(eq => !string.IsNullOrEmpty(
                        GetParamStringValueWithTypeFallback(eq, Config.SourceEquipmentParam)))
                    .ToList();

                UI?.SetStatus($"DBG Step 1b: {candidateEquipment.Count}/{allEquipment.Count} equipment have " +
                    $"non-empty '{Config.SourceEquipmentParam}'.");

                if (candidateEquipment.Count == 0)
                {
                    UI?.SetStatus($"DBG Error: No equipment has a value for param '{Config.SourceEquipmentParam}'. " +
                        "Check the source parameter name in ⚙ Config.");
                    return;
                }

                // ── 2. Transaction: write equipment param to each point ────────
                using (Transaction trans = new Transaction(doc, "HMV - Electrical Flow Inject"))
                {
                    trans.Start();

                    FailureHandlingOptions fho = trans.GetFailureHandlingOptions();
                    fho.SetFailuresPreprocessor(new WarningSuppressor());
                    trans.SetFailureHandlingOptions(fho);

                    int updated   = 0;
                    int skipped   = 0;
                    int noParam   = 0;
                    string lastDebug = string.Empty;

                    foreach (ElementId id in PointIds)
                    {
                        FamilyInstance fi = doc.GetElement(id) as FamilyInstance;
                        if (fi == null) { skipped++; continue; }

                        XYZ            pos = GetAdaptivePosition(doc, fi);
                        FamilyInstance eq  = FindClosestEquipment(candidateEquipment, pos);

                        if (eq == null) { skipped++; continue; }

                        string val = GetParamStringValueWithTypeFallback(eq, Config.SourceEquipmentParam);

                        Parameter dest = fi.LookupParameter(Config.DestPointParam);
                        if (dest == null)
                        {
                            noParam++;
                            lastDebug = $"DBG: Point '{fi.Name}' (id {id.IntegerValue}) — " +
                                $"dest param '{Config.DestPointParam}' NOT FOUND.";
                            continue;
                        }
                        if (dest.IsReadOnly)
                        {
                            noParam++;
                            lastDebug = $"DBG: Point '{fi.Name}' — dest param is READ-ONLY.";
                            continue;
                        }
                        if (dest.StorageType != StorageType.String)
                        {
                            noParam++;
                            lastDebug = $"DBG: Point '{fi.Name}' — dest param StorageType={dest.StorageType} (not String).";
                            continue;
                        }

                        dest.Set(val ?? string.Empty);
                        lastDebug = $"DBG: Point '{fi.Name}' ← '{val}' (from eq '{eq.Name}')";
                        updated++;
                    }

                    trans.Commit();

                    string summary = $"✔ {updated} updated, {skipped} skipped, {noParam} param-errors.";
                    if (noParam > 0 || skipped > 0)
                        summary += " | Last: " + lastDebug;
                    UI?.SetStatus(summary);
                }
            }
            catch (Exception ex)
            {
                UI?.SetStatus("DBG Exception: " + ex.GetType().Name + " — " + ex.Message);
            }
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private static XYZ GetAdaptivePosition(Document doc, FamilyInstance fi)
        {
            try
            {
                IList<ElementId> ptIds =
                    AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(fi);

                if (ptIds != null && ptIds.Count > 0)
                {
                    ReferencePoint rp = doc.GetElement(ptIds[0]) as ReferencePoint;
                    if (rp != null) return rp.Position;
                }
            }
            catch { }

            return (fi.Location as LocationPoint)?.Point ?? XYZ.Zero;
        }

        private static FamilyInstance FindClosestEquipment(
            List<FamilyInstance> candidates, XYZ position)
        {
            FamilyInstance closest = null;
            double minDist = double.MaxValue;

            foreach (FamilyInstance eq in candidates)
            {
                XYZ loc = (eq.Location as LocationPoint)?.Point;
                if (loc == null) continue;

                double dist = position.DistanceTo(loc);
                if (dist < minDist) { minDist = dist; closest = eq; }
            }

            return closest;
        }

        private static string GetParamStringValueWithTypeFallback(FamilyInstance fi, string paramName)
        {
            if (fi == null || string.IsNullOrWhiteSpace(paramName)) return string.Empty;

            Parameter p = fi.LookupParameter(paramName);

            bool instanceEmpty = p == null
                || !p.HasValue
                || (p.StorageType == StorageType.String && string.IsNullOrEmpty(p.AsString()));

            if (instanceEmpty && fi.Symbol != null)
            {
                foreach (Parameter sp in fi.Symbol.Parameters)
                {
                    if (sp.Definition.Name == paramName) { p = sp; break; }
                }
            }

            if (p == null) return string.Empty;

            switch (p.StorageType)
            {
                case StorageType.String:  return p.AsString() ?? string.Empty;
                case StorageType.Integer: return p.AsInteger().ToString(Inv);
                case StorageType.Double:  return p.AsDouble().ToString(Inv);
                default:                  return string.Empty;
            }
        }

        private static void SetStringParam(FamilyInstance fi, string paramName, string value)
        {
            if (fi == null || string.IsNullOrWhiteSpace(paramName)) return;
            Parameter p = fi.LookupParameter(paramName);
            if (p == null || p.IsReadOnly || p.StorageType != StorageType.String) return;
            p.Set(value ?? string.Empty);
        }
    }
}
