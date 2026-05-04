using Autodesk.Revit.UI;
using System;
using System.Windows;
using System.Windows.Input;

namespace HMVTools
{
    public partial class ElectricalSuiteWindow : Window
    {
        private readonly UIApplication _uiapp;

        // ── Sub-tool state (lazy) ─────────────────────────────────────────
        private ElectricalFlowPickHandler _flowPickHandler;
        private ExternalEvent _flowPickEvent;
        private ElectricalFlowExportHandler _flowExportHandler;
        private ExternalEvent _flowExportEvent;
        private ElectricalFlowWindow _flowWindow;

        private ElectricalRefresherPickHandler _refreshPickHandler;
        private ExternalEvent                  _refreshPickEvent;
        private ElectricalRefresherHandler     _refreshRunHandler;
        private ExternalEvent                  _refreshRunEvent;
        private ElectricalRefresherWindow      _refreshWindow;

        public ElectricalSuiteWindow(UIApplication uiapp)
        {
            InitializeComponent();
            _uiapp = uiapp;

            // ExternalEvent.Create MUST be called in a valid Revit API context.
            // The constructor is called from IExternalCommand.Execute, so it is safe here.
            _flowPickHandler   = new ElectricalFlowPickHandler();
            _flowPickEvent     = ExternalEvent.Create(_flowPickHandler);
            _flowExportHandler = new ElectricalFlowExportHandler();
            _flowExportEvent   = ExternalEvent.Create(_flowExportHandler);

            _refreshPickHandler = new ElectricalRefresherPickHandler();
            _refreshPickEvent   = ExternalEvent.Create(_refreshPickHandler);
            _refreshRunHandler  = new ElectricalRefresherHandler();
            _refreshRunEvent    = ExternalEvent.Create(_refreshRunHandler);

            this.Closed += (s, e) => ElectricalSuiteCommand.ClearWindow();
        }

        // ═══════════════════════════════════════════════════════════════════
        //  BUTTON HANDLERS
        // ═══════════════════════════════════════════════════════════════════

        private void BtnFlow_Click(object sender, RoutedEventArgs e)
        {
            if (_flowWindow != null && _flowWindow.IsLoaded)
            {
                _flowWindow.Focus();
                return;
            }

            _flowWindow = new ElectricalFlowWindow(
                _uiapp,
                _flowPickHandler,  _flowPickEvent,
                _flowExportHandler, _flowExportEvent);

            var helper = new System.Windows.Interop.WindowInteropHelper(_flowWindow);
            helper.Owner = new System.Windows.Interop.WindowInteropHelper(this).Handle;

            _flowWindow.Closed += (s, ev) =>
            {
                _flowWindow = null;
                this.Show();
                this.Activate();
                SetStatus("Heredate Parameter closed.");
            };

            SetStatus("Opening Heredate Parameter...");
            this.Hide();
            _flowWindow.Show();
        }

        private void BtnGeom_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var win = new ElectricalGeometryWindow();
                new System.Windows.Interop.WindowInteropHelper(win)
                {
                    Owner = new System.Windows.Interop.WindowInteropHelper(this).Handle
                };

                SetStatus("Opening Export Information...");
                win.ShowDialog();
                SetStatus("Export Information closed.");
            }
            catch (Exception ex)
            {
                SetStatus("Error: " + ex.Message);
            }
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            if (_refreshWindow != null && _refreshWindow.IsLoaded)
            {
                _refreshWindow.Focus();
                return;
            }

            _refreshWindow = new ElectricalRefresherWindow(
                _uiapp,
                _refreshPickHandler, _refreshPickEvent,
                _refreshRunHandler,  _refreshRunEvent);

            var helper = new System.Windows.Interop.WindowInteropHelper(_refreshWindow);
            helper.Owner = new System.Windows.Interop.WindowInteropHelper(this).Handle;

            _refreshWindow.Closed += (s, ev) =>
            {
                _refreshWindow = null;
                this.Show();
                this.Activate();
                SetStatus("Refresh Connections closed.");
            };

            SetStatus("Opening Refresh Connections...");
            this.Hide();
            _refreshWindow.Show();
        }

        // ═══════════════════════════════════════════════════════════════════
        //  WINDOW CHROME
        // ═══════════════════════════════════════════════════════════════════

        private void BtnClose_Click(object sender, RoutedEventArgs e) => this.Close();

        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => this.DragMove();

        private void SetStatus(string msg)
        {
            Dispatcher.BeginInvoke(new Action(() => txtStatus.Text = msg));
        }
    }
}
