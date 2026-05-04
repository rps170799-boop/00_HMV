using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace HMVTools
{
    public partial class ElectricalFlowWindow : Window
    {
        private readonly UIApplication              _uiapp;
        private readonly ElectricalFlowPickHandler   _pickHandler;
        private readonly ExternalEvent               _pickEvent;
        private readonly ElectricalFlowExportHandler _exportHandler;
        private readonly ExternalEvent               _exportEvent;

        private List<ElementId>      _pointIds = new List<ElementId>();
        private ElectricalFlowConfig _config   = new ElectricalFlowConfig();

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

            this.Closed += (s, e) => ElectricalFlowCommand.ClearWindow();
        }

        // ═══════════════════════════════════════════════════════════════════
        //  CALLED BY PICK HANDLER
        // ═══════════════════════════════════════════════════════════════════

        public void OnPointsPicked(List<ElementId> ids)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                _pointIds = ids;
                txtPointCount.Text = $"{ids.Count} point(s) selected";
                SetStatus($"{ids.Count} adaptive point(s) ready.");
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
                txtStatus.Text += message + Environment.NewLine;
                txtStatus.ScrollToEnd();
            }));
        }

        public void ClearLog()
        {
            Dispatcher.BeginInvoke(new Action(() => txtStatus.Text = string.Empty));
        }

        // ═══════════════════════════════════════════════════════════════════
        //  BUTTON HANDLERS
        // ═══════════════════════════════════════════════════════════════════

        private void BtnPick_Click(object sender, RoutedEventArgs e)
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
                .FirstOrDefault(fi => AdaptiveComponentInstanceUtils.IsAdaptiveComponentInstance(fi));

            FamilyInstance sampleEquipment = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .FirstOrDefault();

            List<string> equipParams = BuildParamListAllLevels(sampleEquipment);
            List<string> pointParams = BuildWritableInstanceParamList(samplePoint);

            if (equipParams.Count == 0) equipParams.Add("(no equipment found in model)");
            if (pointParams.Count == 0) pointParams.Add("(no writable instance params found)");

            var configWin = new ElectricalFlowConfigWindow(equipParams, pointParams, _config);
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

        private void BtnClearPoints_Click(object sender, RoutedEventArgs e)
        {
            _pointIds.Clear();
            txtPointCount.Text = "No points selected";
            SetStatus("Selection cleared.");
        }

        private void BtnInject_Click(object sender, RoutedEventArgs e)
        {
            if (_pointIds.Count == 0)
            {
                SetStatus("Error: Select adaptive points first.");
                return;
            }

            if (string.IsNullOrWhiteSpace(_config.SourceEquipmentParam))
            {
                SetStatus("Error: Open ⚙ Config and set the parameter names first.");
                return;
            }

            _exportHandler.PointIds = _pointIds;
            _exportHandler.Config   = _config;
            _exportHandler.UI       = this;

            ClearLog();
            SetStatus("Executing — please wait...");
            _exportEvent.Raise();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e) => this.Close();

        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => this.DragMove();

        // ═══════════════════════════════════════════════════════════════════
        //  HELPERS
        // ═══════════════════════════════════════════════════════════════════

        private static List<string> BuildParamListAllLevels(FamilyInstance fi)
        {
            if (fi == null) return new List<string>();

            var instanceNames = fi.Parameters.Cast<Parameter>()
                .Where(p => p.StorageType != StorageType.ElementId
                         && p.StorageType != StorageType.None
                         && !string.IsNullOrWhiteSpace(p.Definition.Name))
                .Select(p => p.Definition.Name);

            var typeNames = fi.Symbol != null
                ? fi.Symbol.Parameters.Cast<Parameter>()
                    .Where(p => p.StorageType != StorageType.ElementId
                             && p.StorageType != StorageType.None
                             && !string.IsNullOrWhiteSpace(p.Definition.Name))
                    .Select(p => p.Definition.Name)
                : Enumerable.Empty<string>();

            return instanceNames.Concat(typeNames).Distinct().OrderBy(n => n).ToList();
        }

        private static List<string> BuildWritableInstanceParamList(FamilyInstance fi)
        {
            if (fi == null) return new List<string>();

            return fi.Parameters.Cast<Parameter>()
                .Where(p => !p.IsReadOnly
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
