using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using static HMVTools.ElectricalConnectionCommand;

namespace HMVTools
{
    // ─── Dedicated handler for picking snap points ─────────────────────────────
    public class PickPointHandler : IExternalEventHandler
    {
        public enum PickTarget { A, B }

        public PickTarget Target   { get; set; }
        public ConexionFlexPipeWindow UI { get; set; }

        /// <summary>
        /// When true, skips the face-step and picks an adaptive component element,
        /// using its first adaptive point as the picked XYZ.
        /// </summary>
        public bool UseAdaptivePick { get; set; }

        public XYZ PickedPointA { get; private set; }
        public XYZ PickedPointB { get; private set; }

        public string GetName() => "PickPointHandler";

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document   doc   = uidoc.Document;

            try
            {
                string label = Target == PickTarget.A ? "Point A (Start)" : "Point B (End)";

                XYZ picked;

                if (UseAdaptivePick)
                {
                    // ── Adaptive component path ────────────────────────────────────────
                    // Single click on the analytic family instance; reads its first
                    // adaptive reference point position automatically.
                    Reference elemRef = uidoc.Selection.PickObject(
                        ObjectType.Element,
                        new AdaptiveComponentFilter(),
                        $"Select Adaptive Component for {label} — ESC to cancel");

                    FamilyInstance fi = doc.GetElement(elemRef) as FamilyInstance;
                    if (fi == null)
                        throw new InvalidOperationException("Selected element is not a FamilyInstance.");

                    IList<ElementId> pointIds =
                        AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(fi);

                    if (pointIds == null || pointIds.Count == 0)
                        throw new InvalidOperationException(
                            "The selected adaptive component has no placement points.");

                    // Take the first reference point as the representative location
                    ReferencePoint rp = doc.GetElement(pointIds[0]) as ReferencePoint;
                    if (rp == null)
                        throw new InvalidOperationException(
                            "Could not retrieve the ReferencePoint from the adaptive component.");

                    picked = rp.Position;
                }
                else
                {
                    // ── Original face → snap point path (unchanged) ───────────────────
                    Reference faceRef = uidoc.Selection.PickObject(
                        ObjectType.Face,
                        new MEPElementFilter(),
                        $"Step 1: Select the flat FACE of the equipment for {label} — ESC to cancel");

                    using (Transaction t = new Transaction(doc, "Set Temp Work Plane"))
                    {
                        t.Start();
                        SketchPlane sp = SketchPlane.Create(doc, faceRef);
                        uidoc.ActiveView.SketchPlane = sp;
                        t.Commit();
                    }

                    ObjectSnapTypes snaps =
                        ObjectSnapTypes.Midpoints  |
                        ObjectSnapTypes.Centers    |
                        ObjectSnapTypes.Endpoints  |
                        ObjectSnapTypes.Intersections |
                        ObjectSnapTypes.Nearest;

                    picked = uidoc.Selection.PickPoint(
                        snaps,
                        $"Step 2: Now use Snaps (Midpoint, Center, Endpoint) on the selected face");
                }

                if (Target == PickTarget.A) PickedPointA = picked;
                else                        PickedPointB = picked;

                UI?.OnPointPicked(Target, picked);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                UI?.SetStatus("Selection canceled by user.");
                UI?.RestoreWindow();
            }
            catch (Exception ex)
            {
                UI?.SetStatus($"Error capturing point: {ex.Message}");
                UI?.RestoreWindow();
            }
        }
    }

    /// <summary>
    /// Restricts selection to FamilyInstance elements that are adaptive components
    /// (i.e. have at least one adaptive point — "Reference Point" placement).
    /// </summary>
    public class AdaptiveComponentFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is FamilyInstance fi
                && AdaptiveComponentInstanceUtils.IsAdaptiveComponentInstance(fi);
        }

        public bool AllowReference(Reference reference, XYZ position) => false;
    }

    /// <summary>
    /// Restricts selection to FamilyInstance elements only (equipment, fixtures).
    /// Rejects grids, levels, annotations so the cursor snaps cleanly to geometry.
    /// </summary>
    public class MEPElementFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            // Accept any FamilyInstance (electrical equipment, generic models, etc.)
            return elem is FamilyInstance;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            // Accept all geometry faces/edges on allowed elements
            return true;
        }
    }

    public partial class ConexionFlexPipeWindow : Window
    {
        private UIApplication _uiapp;
        private ElectricalConnectionHandler _handler;
        private ExternalEvent _exEvent;

        // ── Dedicated pick infrastructure ─────────────────────────────────────
        private PickPointHandler _pickHandler;
        private ExternalEvent _pickEvent;

        private XYZ _pointA = null;
        private XYZ _pointB = null;

        public ConexionFlexPipeWindow(UIApplication uiapp, ElectricalConnectionHandler handler, ExternalEvent exEvent)
        {
            InitializeComponent();
            _uiapp = uiapp;
            _handler = handler;
            _exEvent = exEvent;
            _handler.UI = this;

            // Create dedicated pick handler + event
            _pickHandler = new PickPointHandler { UI = this };
            _pickEvent = ExternalEvent.Create(_pickHandler);

            this.Closed += (s, e) => { ElectricalConnectionCommand.ClearWindow(); };

            LoadRevitTypes();
        }

        // Called by PickPointHandler.Execute() after a successful pick
        public void OnPointPicked(PickPointHandler.PickTarget target, XYZ point)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (target == PickPointHandler.PickTarget.A)
                {
                    _pointA = point;
                    lblElementA.Text = $"Point A: X:{point.X:F3}  Y:{point.Y:F3}  Z:{point.Z:F3}";
                    lblElementA.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(0, 130, 60)); // green = confirmed
                    txtStatus.Text = "Point A selected ✔";
                }
                else
                {
                    _pointB = point;
                    lblElementB.Text = $"Point B: X:{point.X:F3}  Y:{point.Y:F3}  Z:{point.Z:F3}";
                    lblElementB.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(0, 130, 60)); // green = confirmed
                    txtStatus.Text = "Point B selected ✔";
                }
                this.Show();
                this.Activate();
            }));
        }

        // Called by handler when user presses ESC
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

        // ─── Button clicks: hide window THEN raise the external event ─────────

        private void BtnPickA_Click(object sender, RoutedEventArgs e)
        {
            _pickHandler.Target          = PickPointHandler.PickTarget.A;
            _pickHandler.UseAdaptivePick = false;   // ← face workflow
            this.Hide();
            _pickEvent.Raise();
        }

        private void BtnPickB_Click(object sender, RoutedEventArgs e)
        {
            _pickHandler.Target          = PickPointHandler.PickTarget.B;
            _pickHandler.UseAdaptivePick = false;   // ← face workflow
            this.Hide();
            _pickEvent.Raise();
        }

        // ─── Adaptive component pick (analytic family / reference point) ──────

        private void BtnPickAdaptiveA_Click(object sender, RoutedEventArgs e)
        {
            _pickHandler.Target         = PickPointHandler.PickTarget.A;
            _pickHandler.UseAdaptivePick = true;
            this.Hide();
            _pickEvent.Raise();
        }

        private void BtnPickAdaptiveB_Click(object sender, RoutedEventArgs e)
        {
            _pickHandler.Target         = PickPointHandler.PickTarget.B;
            _pickHandler.UseAdaptivePick = true;
            this.Hide();
            _pickEvent.Raise();
        }

        // ─── Rest of the UI ───────────────────────────────────────────────────

        private void LoadRevitTypes()
        {
            Document doc = _uiapp.ActiveUIDocument.Document;

            var flexTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(FlexPipeType))
                .Cast<FlexPipeType>()
                .OrderBy(t => t.Name)
                .ToList();
            cmbFlexPipeType.ItemsSource = flexTypes;
            if (flexTypes.Any()) cmbFlexPipeType.SelectedIndex = 0;

            var sysTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(MEPSystemType))
                .Cast<MEPSystemType>()
                .OrderBy(t => t.Name)
                .ToList();

            cmbSystemType.ItemsSource = sysTypes;
            if (cmbSystemType.Items.Count > 0) cmbSystemType.SelectedIndex = 0;
        }

        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "DXF Files (*.dxf)|*.dxf",
                Title = "Seleccionar trayectoria DXF"
            };
                if (ofd.ShowDialog() == true)
                txtDxfPath.Text = ofd.FileName;
        }

        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtDxfPath.Text))
            { SetStatus("Error: Seleccione un archivo DXF."); return; }

            if (_pointA == null || _pointB == null)
            { SetStatus("Error: Debe seleccionar ambos puntos (A y B)."); return; }

            if (cmbFlexPipeType.SelectedItem == null || cmbSystemType.SelectedItem == null)
            { SetStatus("Error: Seleccione los tipos de FlexPipe y Sistema."); return; }

            if (string.IsNullOrWhiteSpace(txtParamName.Text))
            { SetStatus("Error: Indique el nombre del parámetro compartido."); return; }

            _handler.DxfFilePath = txtDxfPath.Text;
            _handler.PointA = _pointA;
            _handler.PointB = _pointB;
            _handler.SharedParameterName = txtParamName.Text.Trim();
            _handler.SelectedFlexPipeTypeKey = ((ElementType)cmbFlexPipeType.SelectedItem).Name;
            _handler.SelectedElectricalSystemTypeKey = ((ElementType)cmbSystemType.SelectedItem).Name;

            SetStatus("Procesando geometría DXF y generando FlexPipe...");
            _exEvent.Raise();
        }
    }
}