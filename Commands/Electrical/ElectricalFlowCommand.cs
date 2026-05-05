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

                UI?.SetStatus($"DBG Step 1b: {candidateEquipment.Count}/{allEquipment.Count} candidates found.");

                if (candidateEquipment.Count == 0)
                {
                    UI?.SetStatus($"DBG Error: No equipment has '{Config.SourceEquipmentParam}'. Check ⚙ Config.");
                    return;
                }

                // ── 1c. Pre-cache bounding-box centers (avoid repeated get_BoundingBox) ──
                var candidateCache = candidateEquipment
                .Select(eq =>
                {
                    BoundingBoxXYZ bb = null;
                    try { bb = eq.get_BoundingBox(null); } catch { }
                    if (bb == null) return new EquipmentCache { Eq = eq, Center = null };
                    return new EquipmentCache
                    {
                        Eq = eq,
                        Center = (bb.Min + bb.Max) * 0.5,
                        BbMin = bb.Min,
                        BbMax = bb.Max
                    };
                })
                .Where(x => x.Center != null)
                .ToList();

                UI?.SetStatus($"DBG Step 1c: {candidateCache.Count} candidates with valid locations.");

                // Build a one-line candidate summary visible after the run
                string candidateSummary = string.Join(" | ", candidateCache.Select(x =>
                {
                    string v = GetParamStringValueWithTypeFallback(x.Eq, Config.SourceEquipmentParam);
                    string loc = $"({x.Center.X:F1},{x.Center.Y:F1},{x.Center.Z:F1})";
                    return $"'{x.Eq.Name}'(id:{x.Eq.Id.IntegerValue} val:'{v}' ctr:{loc})";
                }));

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
                        FamilyInstance eq = FindClosestEquipment(candidateCache, pos);

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

                        // Inject as "PARAM_VALUE_ELEMENTID" using the actual Revit ElementId.
                        // Example: source param = "TC", eq.Id = 2412545 → "TC_2412545"
                        string valueToInject = string.IsNullOrWhiteSpace(val)
                            ? eq.Id.IntegerValue.ToString()
                            : val + "_" + eq.Id.IntegerValue;

                        dest.Set(valueToInject);
                        lastDebug = $"DBG: Point '{fi.Name}' (id {id.IntegerValue}) ← '{valueToInject}'" +
                                    $" | src param='{Config.SourceEquipmentParam}'" +
                                    $" | dst param='{Config.DestPointParam}'" +
                                    $" | eq='{eq.Name}' (id {eq.Id.IntegerValue})";
                        updated++;
                    }

                    trans.Commit();

                    string summary = $"✔ {updated} updated, {skipped} skipped, {noParam} param-errors. | CANDIDATES: {candidateSummary} | Last: {lastDebug}";
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
    List<EquipmentCache> cachedCandidates, XYZ position)
        {
            if (cachedCandidates.Count == 0) return null;

            const double maxRadiusFt = 50.0; // skip matches farther than this in XY

            FamilyInstance closest = null;
            double minDist = double.MaxValue;

            foreach (var c in cachedCandidates)
            {
                // Clamp point XY to bbox extents → nearest point on bbox edge (2D)
                double nearestX = Math.Max(c.BbMin.X, Math.Min(position.X, c.BbMax.X));
                double nearestY = Math.Max(c.BbMin.Y, Math.Min(position.Y, c.BbMax.Y));

                double dx = position.X - nearestX;
                double dy = position.Y - nearestY;
                double dist2D = Math.Sqrt(dx * dx + dy * dy);

                if (dist2D < minDist && dist2D <= maxRadiusFt)
                {
                    minDist = dist2D;
                    closest = c.Eq;
                }
            }

            return closest;
        }
        public struct EquipmentCache
        {
            public FamilyInstance Eq;
            public XYZ Center;
            public XYZ BbMin;
            public XYZ BbMax;
        }
        
        private static void WalkGeometry(GeometryElement geomElem, XYZ position, ref double minDist)
        {
            foreach (GeometryObject geomObj in geomElem)
            {
                Solid solid = geomObj as Solid;
                if (solid != null && solid.Faces.Size > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        IntersectionResult result = face.Project(position);
                        if (result != null && result.Distance < minDist)
                            minDist = result.Distance;
                    }
                    continue;
                }

                GeometryInstance gi = geomObj as GeometryInstance;
                if (gi != null)
                {
                    GeometryElement nested = gi.GetInstanceGeometry();
                    if (nested != null)
                        WalkGeometry(nested, position, ref minDist);
                }
            }
        }
        private static XYZ GetEquipmentCenter(FamilyInstance eq)
        {
            try
            {
                BoundingBoxXYZ bb = eq.get_BoundingBox(null);
                if (bb != null)
                    return (bb.Min + bb.Max) * 0.5;
            }
            catch { }

            return (eq.Location as LocationPoint)?.Point;
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

        

    }
}
