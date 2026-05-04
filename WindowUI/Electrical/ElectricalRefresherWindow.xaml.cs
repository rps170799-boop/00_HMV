using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace HMVTools
{
    public partial class ElectricalRefresherWindow : Window
    {
        private readonly UIApplication            _uiapp;
        private readonly ElectricalRefresherPickHandler _pickHandler;
        private readonly ExternalEvent                  _pickEvent;
        private readonly ElectricalRefresherHandler     _runHandler;
        private readonly ExternalEvent                  _runEvent;

        private List<ElementId>          _pointIds = new List<ElementId>();
        private ElectricalRefresherConfig _config   = new ElectricalRefresherConfig();

        public ElectricalRefresherWindow(
            UIApplication                   uiapp,
            ElectricalRefresherPickHandler  pickHandler, ExternalEvent pickEvent,
            ElectricalRefresherHandler      runHandler,  ExternalEvent runEvent)
        {
            InitializeComponent();

            _uiapp       = uiapp;
            _pickHandler = pickHandler;
            _pickEvent   = pickEvent;
            _runHandler  = runHandler;
            _runEvent    = runEvent;

            _pickHandler.UI = this;
            _runHandler.UI  = this;

            this.Closed += (s, e) => ElectricalRefresherCommand.ClearWindow();
        }

        // ═══════════════════════════════════════════════════════════════════
        //  CALLED BY HANDLERS
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

        public void SetStatus(string msg)
        {
            Dispatcher.BeginInvoke(new Action(() => txtStatus.Text = msg));
        }

        public void Log(string line)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                txtLog.Text += line + Environment.NewLine;
                txtLog.ScrollToEnd();
            }));
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

        private void BtnClearPoints_Click(object sender, RoutedEventArgs e)
        {
            _pointIds.Clear();
            txtPointCount.Text = "No points selected";
            SetStatus("Selection cleared.");
        }

        private void BtnConfig_Click(object sender, RoutedEventArgs e)
        {
            Document doc = _uiapp.ActiveUIDocument.Document;

            // Sample an adaptive point to get its parameter list
            FamilyInstance samplePoint = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .FirstOrDefault(fi => AdaptiveComponentInstanceUtils.IsAdaptiveComponentInstance(fi));

            // Sample a FlexPipe to get its parameter list
            FlexPipe sampleFlexPipe = new FilteredElementCollector(doc)
                .OfClass(typeof(FlexPipe))
                .Cast<FlexPipe>()
                .FirstOrDefault();

            List<string> pointParams    = BuildParamList(samplePoint);
            List<string> flexPipeParams = BuildFlexPipeParamList(sampleFlexPipe);

            // FlexPipe types: "FamilyName : TypeName"
            List<string> flexPipeTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(FlexPipeType))
                .Cast<FlexPipeType>()
                .Select(ft => ft.FamilyName + " : " + ft.Name)
                .OrderBy(s => s)
                .ToList();

            // Piping system types
            List<string> systemTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(PipingSystemType))
                .Cast<PipingSystemType>()
                .Select(st => st.Name)
                .OrderBy(s => s)
                .ToList();

            if (pointParams.Count == 0)    pointParams.Add("(no adaptive points found in model)");
            if (flexPipeParams.Count == 0) flexPipeParams.Add("(no FlexPipes found in model)");
            if (flexPipeTypes.Count == 0)  flexPipeTypes.Add("(no FlexPipe types found)");
            if (systemTypes.Count == 0)    systemTypes.Add("(no PipingSystemTypes found)");

            var configWin = new ElectricalRefresherConfigWindow(
                pointParams, flexPipeParams, flexPipeTypes, systemTypes, _config);

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

        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            if (_pointIds.Count == 0)
            {
                SetStatus("Error: Select adaptive points first.");
                return;
            }

            if (string.IsNullOrWhiteSpace(_config.FlexPipeTypeKey) ||
                string.IsNullOrWhiteSpace(_config.ElectricalSystemTypeKey) ||
                string.IsNullOrWhiteSpace(_config.DxfFolder))
            {
                SetStatus("Error: Open ⚙ Config and complete all fields first.");
                return;
            }

            txtLog.Text = string.Empty;

            _runHandler.PointIds = _pointIds;
            _runHandler.Config   = _config;
            _runHandler.IsReset  = false;
            _runHandler.Config   = _config;
            _runHandler.IsReset  = false;
            _runHandler.UI       = this;

            SetStatus("Running — please wait...");
            _runEvent.Raise();
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            if (_pointIds.Count == 0)
            {
                SetStatus("Error: Select adaptive points first.");
                return;
            }

            MessageBoxResult confirm = System.Windows.MessageBox.Show(
                "This will DELETE all FlexPipes matching the selected points'\n" +
                "connection numbers and reset their Connected flag.\n\nContinue?",
                "HMV Tools — Reset Confirmation",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (confirm != MessageBoxResult.Yes) return;

            txtLog.Text = string.Empty;

            _runHandler.PointIds = _pointIds;
            _runHandler.Config   = _config;
            _runHandler.IsReset  = true;
            _runHandler.UI       = this;

            SetStatus("Resetting — please wait...");
            _runEvent.Raise();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e) => this.Close();

        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => this.DragMove();

        // ═══════════════════════════════════════════════════════════════════
        //  HELPERS — param list builders
        // ═══════════════════════════════════════════════════════════════════

        private static List<string> BuildParamList(FamilyInstance fi)
        {
            if (fi == null) return new List<string>();

            return fi.Parameters.Cast<Parameter>()
                .Where(p => p.StorageType != StorageType.ElementId
                         && p.StorageType != StorageType.None
                         && !string.IsNullOrWhiteSpace(p.Definition.Name))
                .Select(p => p.Definition.Name)
                .Distinct()
                .OrderBy(n => n)
                .ToList();
        }

        private static List<string> BuildFlexPipeParamList(FlexPipe fp)
        {
            if (fp == null) return new List<string>();

            return fp.Parameters.Cast<Parameter>()
                .Where(p => !p.IsReadOnly
                         && p.StorageType == StorageType.String
                         && !string.IsNullOrWhiteSpace(p.Definition.Name))
                .Select(p => p.Definition.Name)
                .Distinct()
                .OrderBy(n => n)
                .ToList();
        }
    }
}
