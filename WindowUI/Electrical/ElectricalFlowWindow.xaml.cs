using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace HMVTools
{
    public partial class ElectricalFlowWindow : Window
    {
        // ── Infrastructure ────────────────────────────────────────────────────
        private readonly UIApplication              _uiapp;
        private readonly ElectricalFlowPickHandler   _pickHandler;
        private readonly ExternalEvent               _pickEvent;
        private readonly ElectricalFlowExportHandler _exportHandler;
        private readonly ExternalEvent               _exportEvent;

        // ── Observable list bound to the ItemsControl ─────────────────────────
        private readonly ObservableCollection<ElectricalFlowPointGroup> _groups =
            new ObservableCollection<ElectricalFlowPointGroup>();

        // ── Configuration (set via Config window) ─────────────────────────────
        private ElectricalFlowConfig _config = new ElectricalFlowConfig();

        public ElectricalFlowWindow(
            UIApplication uiapp,
            ElectricalFlowPickHandler    pickHandler,   ExternalEvent pickEvent,
            ElectricalFlowExportHandler  exportHandler, ExternalEvent exportEvent)
        {
            InitializeComponent();

            _uiapp         = uiapp;
            _pickHandler   = pickHandler;
            _pickEvent     = pickEvent;
            _exportHandler = exportHandler;
            _exportEvent   = exportEvent;

            _pickHandler.UI   = this;
            _exportHandler.UI = this;

            lstGroups.ItemsSource = _groups;

            this.Closed += (s, e) => ElectricalFlowCommand.ClearWindow();
        }

        // ═══════════════════════════════════════════════════════════════════
        //  CALLED BY PICK HANDLER (background thread → Dispatcher)
        // ═══════════════════════════════════════════════════════════════════

        public void OnGroupPicked(List<ElementId> ids)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                var group = new ElectricalFlowPointGroup
                {
                    GroupIndex = _groups.Count,
                    PointIds   = ids
                };

                // NO PropertyChanged wiring here — direction changes are handled
                // exclusively through Click events to avoid WPF re-entrancy.

                _groups.Add(group);
                SetStatus($"Group {group.GroupIndex + 1} added — {ids.Count} point(s).");

                this.Show();
                this.Activate();
            }));
        }

        public void RestoreWindow()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                this.Show();
                this.Activate();
            }));
        }

        public void SetStatus(string message)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                txtStatus.Text = message;
            }));
        }

        // ═══════════════════════════════════════════════════════════════════
        //  SORT DIRECTION CLICK HANDLERS
        //  Click fires only on real user interaction — never on programmatic
        //  IsChecked changes — which is exactly what we need to avoid
        //  the re-entrancy loop that caused the fatal crash.
        // ═══════════════════════════════════════════════════════════════════

        private void ChkSortRight_Click(object sender, RoutedEventArgs e)
        {
            if (!((sender as CheckBox)?.Tag is ElectricalFlowPointGroup g)) return;
            bool nowChecked = ((CheckBox)sender).IsChecked == true;
            g.SortX = nowChecked ? SortXDirection.Right : SortXDirection.None;
            // Manually push the sibling checkbox state (OneWay binding won't re-enter)
            g.OnPropertyChanged(nameof(ElectricalFlowPointGroup.SortLeft));
        }

        private void ChkSortLeft_Click(object sender, RoutedEventArgs e)
        {
            if (!((sender as CheckBox)?.Tag is ElectricalFlowPointGroup g)) return;
            bool nowChecked = ((CheckBox)sender).IsChecked == true;
            g.SortX = nowChecked ? SortXDirection.Left : SortXDirection.None;
            g.OnPropertyChanged(nameof(ElectricalFlowPointGroup.SortRight));
        }

        private void ChkSortUp_Click(object sender, RoutedEventArgs e)
        {
            if (!((sender as CheckBox)?.Tag is ElectricalFlowPointGroup g)) return;
            bool nowChecked = ((CheckBox)sender).IsChecked == true;
            g.SortY = nowChecked ? SortYDirection.Up : SortYDirection.None;
            g.OnPropertyChanged(nameof(ElectricalFlowPointGroup.SortDown));
        }

        private void ChkSortDown_Click(object sender, RoutedEventArgs e)
        {
            if (!((sender as CheckBox)?.Tag is ElectricalFlowPointGroup g)) return;
            bool nowChecked = ((CheckBox)sender).IsChecked == true;
            g.SortY = nowChecked ? SortYDirection.Down : SortYDirection.None;
            g.OnPropertyChanged(nameof(ElectricalFlowPointGroup.SortUp));
        }

        // ═══════════════════════════════════════════════════════════════════
        //  BUTTON HANDLERS
        // ═══════════════════════════════════════════════════════════════════

        private void BtnAddGroup_Click(object sender, RoutedEventArgs e)
        {
            _pickHandler.UI = this;
            this.Hide();
            _pickEvent.Raise();
        }

        private void BtnConfig_Click(object sender, RoutedEventArgs e)
        {
            Document doc = _uiapp.ActiveUIDocument.Document;

            FamilyInstance samplePoint = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .FirstOrDefault(fi =>
                    AdaptiveComponentInstanceUtils.IsAdaptiveComponentInstance(fi));

            FamilyInstance sampleEquipment = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .FirstOrDefault();

            // Dropdown 1 — source on equipment: instance + type params
            List<string> equipParams = BuildParamListAllLevels(sampleEquipment);

            // Dropdowns 2 & 3 — destination on adaptive point:
            // ONLY shared, writable, INSTANCE-level parameters.
            // Type params on adaptive components cannot be set per-instance,
            // so they must not appear here.
            List<string> pointSharedInstanceParams = BuildSharedInstanceParamList(samplePoint);

            if (equipParams.Count == 0)
                equipParams.Add("(no equipment found in model)");
            if (pointSharedInstanceParams.Count == 0)
                pointSharedInstanceParams.Add("(no shared instance params found on adaptive component)");

            var configWin = new ElectricalFlowConfigWindow(
                equipParams, pointSharedInstanceParams, _config);

            new System.Windows.Interop.WindowInteropHelper(configWin)
            {
                Owner = new System.Windows.Interop.WindowInteropHelper(this).Handle
            };

            if (configWin.ShowDialog() == true)
            {
                _config = configWin.Result;
                SetStatus("Configuration saved.");
            }
        }

        private void BtnRemoveGroup_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as Button)?.Tag is ElectricalFlowPointGroup group)
            {
                _groups.Remove(group);
                ReindexGroups();
                SetStatus("Group removed.");
            }
        }

        private void BtnClearAll_Click(object sender, RoutedEventArgs e)
        {
            _groups.Clear();
            SetStatus("All groups cleared.");
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (_groups.Count == 0)
            {
                SetStatus("Error: Add at least one group before exporting.");
                return;
            }

            if (!_groups.Any(g => g.PointIds.Count >= 2))
            {
                SetStatus("Error: Each group needs at least 2 points.");
                return;
            }

            if (string.IsNullOrWhiteSpace(_config.SourceEquipmentParam))
            {
                SetStatus("Error: Open ⚙ Config and set the parameter names first.");
                return;
            }

            _exportHandler.Groups = _groups.ToList();
            _exportHandler.Config = _config;
            _exportHandler.UI     = this;

            SetStatus("Executing — please wait...");
            _exportEvent.Raise();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e) => this.Close();

        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }

        // ═══════════════════════════════════════════════════════════════════
        //  HELPERS
        // ═══════════════════════════════════════════════════════════════════

        private void ReindexGroups()
        {
            for (int i = 0; i < _groups.Count; i++)
            {
                _groups[i].GroupIndex = i;
                _groups[i].OnPropertyChanged(nameof(ElectricalFlowPointGroup.DisplayLabel));
            }
        }

        /// <summary>
        /// Dropdown 1 — equipment source: all readable params (instance + type).
        /// </summary>
        private static List<string> BuildParamListAllLevels(FamilyInstance fi)
        {
            if (fi == null) return new List<string>();

            var instanceNames = fi.Parameters
                .Cast<Parameter>()
                .Where(p => p.StorageType != StorageType.ElementId
                         && p.StorageType != StorageType.None
                         && !string.IsNullOrWhiteSpace(p.Definition.Name))
                .Select(p => p.Definition.Name);

            var typeNames = fi.Symbol != null
                ? fi.Symbol.Parameters
                    .Cast<Parameter>()
                    .Where(p => p.StorageType != StorageType.ElementId
                             && p.StorageType != StorageType.None
                             && !string.IsNullOrWhiteSpace(p.Definition.Name))
                    .Select(p => p.Definition.Name)
                : Enumerable.Empty<string>();

            return instanceNames
                .Concat(typeNames)
                .Distinct()
                .OrderBy(n => n)
                .ToList();
        }

        /// <summary>
        /// Dropdowns 2 & 3 — adaptive point destination: ALL writable instance-level
        /// parameters (shared or non-shared). Type parameters are excluded because
        /// they cannot be set per-instance from outside the family editor.
        /// </summary>
        private static List<string> BuildSharedInstanceParamList(FamilyInstance fi)
        {
            if (fi == null) return new List<string>();

            // fi.Parameters contains only INSTANCE-bound parameters for this element.
            // Symbol.Parameters contains the TYPE-bound ones — we deliberately skip those.
            return fi.Parameters
                .Cast<Parameter>()
                .Where(p =>
                    !p.IsReadOnly
                    && p.StorageType != StorageType.ElementId
                    && p.StorageType != StorageType.None
                    && !string.IsNullOrWhiteSpace(p.Definition.Name))
                .Select(p => p.Definition.Name)
                .Distinct()
                .OrderBy(n => n)
                .ToList();
        }
    }
}
