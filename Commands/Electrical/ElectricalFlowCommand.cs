using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;

namespace HMVTools
{
    // ═══════════════════════════════════════════════════════════════════
    //  ENUMS
    // ═══════════════════════════════════════════════════════════════════

    public enum SortXDirection { None, Right, Left }
    public enum SortYDirection { None, Up, Down }

    // ═══════════════════════════════════════════════════════════════════
    //  DATA MODEL  —  one row in the main window list
    // ═══════════════════════════════════════════════════════════════════

    public class ElectricalFlowPointGroup : INotifyPropertyChanged
    {
        private SortXDirection _sortX = SortXDirection.None;
        private SortYDirection _sortY = SortYDirection.None;
        private bool _linkedToNext;

        public int GroupIndex { get; set; }
        public List<ElementId> PointIds { get; set; } = new List<ElementId>();

        // ── X axis (Right / Left are mutually exclusive) ─────────────────────
        public SortXDirection SortX
        {
            get => _sortX;
            set
            {
                if (_sortX == value) return;
                _sortX = value;
                OnPropertyChanged(nameof(SortX));
                OnPropertyChanged(nameof(SortRight));
                OnPropertyChanged(nameof(SortLeft));
            }
        }

        public bool SortRight
        {
            get => SortX == SortXDirection.Right;
            set
            {
                if (value) SortX = SortXDirection.Right;
                else if (SortX == SortXDirection.Right) SortX = SortXDirection.None;
            }
        }

        public bool SortLeft
        {
            get => SortX == SortXDirection.Left;
            set
            {
                if (value) SortX = SortXDirection.Left;
                else if (SortX == SortXDirection.Left) SortX = SortXDirection.None;
            }
        }

        // ── Y axis (Up / Down are mutually exclusive) ─────────────────────────
        public SortYDirection SortY
        {
            get => _sortY;
            set
            {
                if (_sortY == value) return;
                _sortY = value;
                OnPropertyChanged(nameof(SortY));
                OnPropertyChanged(nameof(SortUp));
                OnPropertyChanged(nameof(SortDown));
            }
        }

        public bool SortUp
        {
            get => SortY == SortYDirection.Up;
            set
            {
                if (value) SortY = SortYDirection.Up;
                else if (SortY == SortYDirection.Up) SortY = SortYDirection.None;
            }
        }

        public bool SortDown
        {
            get => SortY == SortYDirection.Down;
            set
            {
                if (value) SortY = SortYDirection.Down;
                else if (SortY == SortYDirection.Down) SortY = SortYDirection.None;
            }
        }

        // ── Padlock: last point of this group links to first of next ──────────
        public bool LinkedToNext
        {
            get => _linkedToNext;
            set { _linkedToNext = value; OnPropertyChanged(nameof(LinkedToNext)); }
        }

        public string DisplayLabel => $"Group {GroupIndex + 1}   —   {PointIds.Count} pts";

        public event PropertyChangedEventHandler PropertyChanged;
        internal void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // ═══════════════════════════════════════════════════════════════════
    //  CONFIGURATION — set by the Config window, consumed by the handler
    // ═══════════════════════════════════════════════════════════════════

    public class ElectricalFlowConfig
    {
        public string SourceEquipmentParam   { get; set; }
        public string DestEquipmentNameParam { get; set; }
        public string DestCNParam            { get; set; }
        public bool   UpdateEquipmentNames   { get; set; } = true;
        public bool   UpdateCN               { get; set; } = true;
    }

    // ═══════════════════════════════════════════════════════════════════
    //  COMMAND  (non-modal pattern — identical to ElectricalConnectionCommand)
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
    //                  returns control to the UI via OnGroupPicked()
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
                    "Select Adaptive Components for this group — click Finish when done");

                if (refs == null || refs.Count == 0)
                {
                    UI?.SetStatus("No elements selected.");
                    UI?.RestoreWindow();
                    return;
                }

                var ids = refs.Select(r => r.ElementId).ToList();
                UI?.OnGroupPicked(ids);
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
    //  EXPORT HANDLER  — sorting, parameter update, CSV generation
    // ═══════════════════════════════════════════════════════════════════

    public class ElectricalFlowExportHandler : IExternalEventHandler
    {
        // Set by the window before raising the event
        public ElectricalFlowWindow              UI     { get; set; }
        public List<ElectricalFlowPointGroup>    Groups { get; set; }
        public ElectricalFlowConfig              Config { get; set; }

        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        public string GetName() => "ElectricalFlowExportHandler";

        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;

            if (Groups == null || Groups.Count == 0 || Config == null)
            {
                UI?.SetStatus("Error: No groups or configuration data available.");
                return;
            }

            try
            {
                UI?.SetStatus("Processing...");

                // ── 1. Collect all Electrical Equipment in the model once ─────
                var allEquipment = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .ToList();

                // Conversion factor: internal Revit feet → millimeters (matches ElectricalGeometryCommand)
                double ftToMm = UnitUtils.ConvertFromInternalUnits(1.0, UnitTypeId.Millimeters);

                // ── 2. Sort each group and store as flat list per group ────────
                var sortedGroups = new List<List<(ElementId Id, XYZ Pos)>>();

                foreach (ElectricalFlowPointGroup group in Groups)
                {
                    var pts = new List<(ElementId Id, XYZ Pos)>();

                    foreach (ElementId id in group.PointIds)
                    {
                        FamilyInstance fi = doc.GetElement(id) as FamilyInstance;
                        if (fi == null) continue;
                        XYZ pos = GetAdaptivePosition(doc, fi);
                        pts.Add((id, pos));
                    }

                    sortedGroups.Add(ApplySort(pts, group.SortX, group.SortY));
                }

                // ── 3. Build connection pairs (global CN counter) ─────────────
                //
                //  Padlock OFF: group pairs are independent  → 1-2-3  then  4-5-6
                //  Padlock ON:  last of current links to first of next → 1-2-3-4-5-6
                //
                var connections = new List<(int CN, ElementId Id1, XYZ Pos1, ElementId Id2, XYZ Pos2)>();
                int cn = 1;

                for (int g = 0; g < sortedGroups.Count; g++)
                {
                    List<(ElementId Id, XYZ Pos)> pts = sortedGroups[g];

                    // Within-group consecutive pairs
                    for (int i = 0; i < pts.Count - 1; i++)
                    {
                        connections.Add((cn++,
                            pts[i].Id,     pts[i].Pos,
                            pts[i + 1].Id, pts[i + 1].Pos));
                    }

                    // Cross-group connection if padlock is ON and a next group exists
                    if (Groups[g].LinkedToNext
                        && g + 1 < sortedGroups.Count
                        && pts.Count > 0
                        && sortedGroups[g + 1].Count > 0)
                    {
                        var next = sortedGroups[g + 1];
                        connections.Add((cn++,
                            pts[pts.Count - 1].Id, pts[pts.Count - 1].Pos,
                            next[0].Id,            next[0].Pos));
                    }
                }

                if (connections.Count == 0)
                {
                    UI?.SetStatus("Error: No connections generated. Each group needs at least 2 points.");
                    return;
                }

                // ── 4. Transaction: write parameters back to adaptive points ──
                //
                // Each CONNECTION is between TWO points (pt1 → pt2).
                // Both endpoints receive the same CN so the connection is
                // fully identified from either end.
                // Equipment names are written per-point: each point gets the
                // name of the equipment closest to it.
                using (Transaction trans = new Transaction(doc, "HMV - Electrical Flow Update"))
                {
                    trans.Start();

                    FailureHandlingOptions fho = trans.GetFailureHandlingOptions();
                    fho.SetFailuresPreprocessor(new WarningSuppressor());
                    trans.SetFailureHandlingOptions(fho);

                    foreach (var conn in connections)
                    {
                        FamilyInstance pt1 = doc.GetElement(conn.Id1) as FamilyInstance;
                        FamilyInstance pt2 = doc.GetElement(conn.Id2) as FamilyInstance;

                        if (Config.UpdateEquipmentNames)
                        {
                            FamilyInstance eq1 = FindClosestEquipment(allEquipment, conn.Pos1);
                            FamilyInstance eq2 = FindClosestEquipment(allEquipment, conn.Pos2);

                            // Read the source value from the equipment — checks instance
                            // param first, then the family type (symbol) param explicitly.
                            string name1 = GetParamStringValueWithTypeFallback(eq1, Config.SourceEquipmentParam);
                            string name2 = GetParamStringValueWithTypeFallback(eq2, Config.SourceEquipmentParam);

                            // Write the equipment name to each endpoint of the connection
                            SetStringParam(pt1, Config.DestEquipmentNameParam, name1);
                            SetStringParam(pt2, Config.DestEquipmentNameParam, name2);
                        }

                        if (Config.UpdateCN)
                        {
                            // CN is written to BOTH endpoints so the connection is
                            // identifiable from either adaptive point.
                            string cnStr = conn.CN.ToString(Inv);
                            SetIntOrStringParam(pt1, Config.DestCNParam, cnStr);
                            SetIntOrStringParam(pt2, Config.DestCNParam, cnStr);
                        }
                    }

                    trans.Commit();
                }

                // ── 5. Build CSV lines ────────────────────────────────────────
                var lines = new List<string>
                {
                    "CN;Equipment 1;Equipment 2;Vano (mm);Desnivel (mm)"
                };

                foreach (var conn in connections)
                {
                    FamilyInstance eq1 = FindClosestEquipment(allEquipment, conn.Pos1);
                    FamilyInstance eq2 = FindClosestEquipment(allEquipment, conn.Pos2);

                    string name1 = GetParamStringValueWithTypeFallback(eq1, Config.SourceEquipmentParam);
                    string name2 = GetParamStringValueWithTypeFallback(eq2, Config.SourceEquipmentParam);

                    // Horizontal distance (vano) and vertical delta (desnivel) in mm
                    // Same formula as ElectricalGeometryCommand: sqrt(dx²+dy²) and abs(dz)
                    double dx = (conn.Pos2.X - conn.Pos1.X) * ftToMm;
                    double dy = (conn.Pos2.Y - conn.Pos1.Y) * ftToMm;
                    double dz = (conn.Pos2.Z - conn.Pos1.Z) * ftToMm;

                    double vano     = Math.Sqrt(dx * dx + dy * dy);
                    double desnivel = Math.Abs(dz);

                    lines.Add(string.Join(";",
                        conn.CN.ToString(Inv),
                        name1,
                        name2,
                        vano.ToString("F4", Inv),
                        desnivel.ToString("F4", Inv)));
                }

                // ── 6. Save file (queued to dispatcher after handler returns) ─
                string statusMsg = $"✔ {connections.Count} connections processed.";
                List<string> csvSnapshot = new List<string>(lines);

                UI?.Dispatcher.BeginInvoke(new Action(() =>
                {
                    var dlg = new Microsoft.Win32.SaveFileDialog
                    {
                        Filter   = "CSV Files (*.csv)|*.csv",
                        Title    = "Save Electrical Flow Report",
                        FileName = "ElectricalFlowReport.csv"
                    };

                    if (dlg.ShowDialog() == true)
                    {
                        File.WriteAllLines(dlg.FileName, csvSnapshot);
                        UI?.SetStatus(statusMsg + " CSV saved.");
                    }
                    else
                    {
                        UI?.SetStatus(statusMsg + " CSV export cancelled.");
                    }
                }));
            }
            catch (Exception ex)
            {
                UI?.SetStatus("Error: " + ex.Message);
            }
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        /// <summary>
        /// Reads the position from the first adaptive reference point of a
        /// FamilyInstance, falling back to LocationPoint if none found.
        /// </summary>
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

        /// <summary>
        /// Applies the selected sort directions to a list of points.
        /// Coordinates are rounded to 3 decimal places (feet) to avoid
        /// Revit floating-point micro-tolerance ordering noise.
        /// </summary>
        private static List<(ElementId Id, XYZ Pos)> ApplySort(
            List<(ElementId Id, XYZ Pos)> pts,
            SortXDirection sortX,
            SortYDirection sortY)
        {
            IEnumerable<(ElementId Id, XYZ Pos)> result = pts;

            // All 4 combinations of X + Y then the two single-axis cases
            if (sortX == SortXDirection.Right && sortY == SortYDirection.Up)
                result = pts.OrderBy(p => Math.Round(p.Pos.X, 3))
                            .ThenBy(p => Math.Round(p.Pos.Y, 3));
            else if (sortX == SortXDirection.Right && sortY == SortYDirection.Down)
                result = pts.OrderBy(p => Math.Round(p.Pos.X, 3))
                            .ThenByDescending(p => Math.Round(p.Pos.Y, 3));
            else if (sortX == SortXDirection.Left && sortY == SortYDirection.Up)
                result = pts.OrderByDescending(p => Math.Round(p.Pos.X, 3))
                            .ThenBy(p => Math.Round(p.Pos.Y, 3));
            else if (sortX == SortXDirection.Left && sortY == SortYDirection.Down)
                result = pts.OrderByDescending(p => Math.Round(p.Pos.X, 3))
                            .ThenByDescending(p => Math.Round(p.Pos.Y, 3));
            else if (sortX == SortXDirection.Right)
                result = pts.OrderBy(p => Math.Round(p.Pos.X, 3));
            else if (sortX == SortXDirection.Left)
                result = pts.OrderByDescending(p => Math.Round(p.Pos.X, 3));
            else if (sortY == SortYDirection.Up)
                result = pts.OrderBy(p => Math.Round(p.Pos.Y, 3));
            else if (sortY == SortYDirection.Down)
                result = pts.OrderByDescending(p => Math.Round(p.Pos.Y, 3));
            // SortX == None && SortY == None → preserve pick order

            return result.ToList();
        }

        /// <summary>
        /// Returns the nearest FamilyInstance from allEquipment to the given position.
        /// Uses LocationPoint for speed (equipment is literally on top of the adaptive point).
        /// </summary>
        private static FamilyInstance FindClosestEquipment(
            List<FamilyInstance> allEquipment, XYZ position)
        {
            FamilyInstance closest = null;
            double minDist = double.MaxValue;

            foreach (FamilyInstance eq in allEquipment)
            {
                XYZ loc = (eq.Location as LocationPoint)?.Point;
                if (loc == null) continue;

                double dist = position.DistanceTo(loc);
                if (dist < minDist)
                {
                    minDist = dist;
                    closest = eq;
                }
            }

            return closest;
        }

        /// <summary>
        /// Reads a parameter value from a FamilyInstance.
        /// Checks the instance parameters first; if not found there, explicitly
        /// walks the family Symbol's parameters (covers shared Type parameters
        /// which LookupParameter on the instance may silently skip when an
        /// instance param of the same name shadows it at a different binding).
        /// </summary>
        private static string GetParamStringValueWithTypeFallback(FamilyInstance fi, string paramName)
        {
            if (fi == null || string.IsNullOrWhiteSpace(paramName)) return string.Empty;

            // 1. Try instance-level parameter first
            Parameter p = fi.LookupParameter(paramName);

            // 2. If not found, has no value, or is an empty string, explicitly search
            //    the Symbol's type parameters — LookupParameter on the instance may
            //    return a shadowing instance param with an empty value instead of the
            //    populated type param of the same name.
            bool instanceEmpty = p == null
                || !p.HasValue
                || (p.StorageType == StorageType.String && string.IsNullOrEmpty(p.AsString()));

            if (instanceEmpty)
            {
                if (fi.Symbol != null)
                {
                    foreach (Parameter sp in fi.Symbol.Parameters)
                    {
                        if (sp.Definition.Name == paramName)
                        {
                            p = sp;
                            break;
                        }
                    }
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

        private static void SetIntOrStringParam(FamilyInstance fi, string paramName, string value)
        {
            if (fi == null || string.IsNullOrWhiteSpace(paramName)) return;
            Parameter p = fi.LookupParameter(paramName);
            if (p == null || p.IsReadOnly) return;

            switch (p.StorageType)
            {
                case StorageType.String:
                    p.Set(value);
                    break;
                case StorageType.Integer:
                    if (int.TryParse(value, out int iv)) p.Set(iv);
                    break;
            }
        }
    }
}
