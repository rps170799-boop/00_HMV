using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Linq;
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

        private ElectricalGeometryHandler _geomHandler;
        private ExternalEvent             _geomEvent;
        private ElectricalGeometryWindow  _geomWindow;

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

            // Geometry handler — must be created here in the valid Revit API context
            var geomDoc = uiapp.ActiveUIDocument.Document;
            var geomParamNames = new System.Collections.Generic.List<string>();
            var geomSeen       = new System.Collections.Generic.HashSet<string>();
            foreach (FlexPipe fp in new Autodesk.Revit.DB.FilteredElementCollector(geomDoc)
                                        .OfClass(typeof(FlexPipe)).Cast<FlexPipe>())
            {
                foreach (Autodesk.Revit.DB.Parameter p in fp.Parameters)
                {
                    try { if (p.Definition != null && !string.IsNullOrWhiteSpace(p.Definition.Name)) geomSeen.Add(p.Definition.Name); } catch { }
                }
            }
            geomParamNames.AddRange(geomSeen);
            geomParamNames.Sort(System.StringComparer.OrdinalIgnoreCase);

            _geomHandler = new ElectricalGeometryHandler(geomDoc);
            _geomEvent   = ExternalEvent.Create(_geomHandler);
            _geomWindow  = new ElectricalGeometryWindow(_geomHandler, _geomEvent, geomParamNames);
            new System.Windows.Interop.WindowInteropHelper(_geomWindow)
            {
                Owner = uiapp.MainWindowHandle
            };
            _geomWindow.Closed += (s, ev) => { _geomWindow = null; };

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
            if (_geomWindow != null && _geomWindow.IsLoaded)
            {
                _geomWindow.Focus();
                return;
            }

            // Recreate window (reuse existing handler/event which were created in the constructor)
            var doc        = _uiapp.ActiveUIDocument.Document;
            var paramNames = new System.Collections.Generic.List<string>();
            var seen       = new System.Collections.Generic.HashSet<string>();
            foreach (FlexPipe fp in new Autodesk.Revit.DB.FilteredElementCollector(doc)
                                        .OfClass(typeof(FlexPipe)).Cast<FlexPipe>())
            {
                foreach (Autodesk.Revit.DB.Parameter p in fp.Parameters)
                {
                    try { if (p.Definition != null && !string.IsNullOrWhiteSpace(p.Definition.Name)) seen.Add(p.Definition.Name); } catch { }
                }
            }
            paramNames.AddRange(seen);
            paramNames.Sort(System.StringComparer.OrdinalIgnoreCase);

            _geomWindow = new ElectricalGeometryWindow(_geomHandler, _geomEvent, paramNames);
            new System.Windows.Interop.WindowInteropHelper(_geomWindow)
            {
                Owner = new System.Windows.Interop.WindowInteropHelper(this).Handle
            };
            _geomWindow.Closed += (s, ev) => { _geomWindow = null; };

            SetStatus("Opening Electrical Geometry...");
            _geomWindow.Show();
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
