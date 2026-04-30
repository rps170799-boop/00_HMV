using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using IxMilia.Dxf;
using IxMilia.Dxf.Entities;
using System;
using System.Collections.Generic;
using System.IO;
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

    /// <summary>
    /// Restricts selection to FlexPipe MEP curves only.
    /// </summary>
    public class FlexPipeSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem) => elem is FlexPipe;
        public bool AllowReference(Reference reference, XYZ position) => false;
    }

    /// <summary>
    /// Picks a FlexPipe and mirrors it about its own A→B axis,
    /// flipping the bulge to the opposite side. Deletes the original.
    /// </summary>
    public class MirrorFlexPipeHandler : IExternalEventHandler
    {
        public ConexionFlexPipeWindow UI { get; set; }

        public string GetName() => "MirrorFlexPipeHandler";

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document   doc   = uidoc.Document;

            try
            {
                // ── 1. Pick the FlexPipe ──────────────────────────────────────
                Reference elemRef = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new FlexPipeSelectionFilter(),
                    "Select the FlexPipe to mirror — ESC to cancel");

                FlexPipe flexPipe = doc.GetElement(elemRef) as FlexPipe;
                if (flexPipe == null)
                    throw new InvalidOperationException("Selected element is not a FlexPipe.");

                // ── 2. Get trajectory points ──────────────────────────────────
                IList<XYZ> pts = flexPipe.Points;
                if (pts == null || pts.Count < 3)
                    throw new InvalidOperationException(
                        "FlexPipe has too few points to determine a mirror plane.");

                XYZ ptA = pts[0];
                XYZ ptB = pts[pts.Count - 1];
                XYZ uAB = (ptB - ptA).Normalize();

                // ── 3. Find the maximum-deviation (bulge) point ───────────────
                // Project each intermediate point onto the A-B line and measure
                // the perpendicular distance; take the largest one.
                XYZ bulgePt  = null;
                double maxDev = 0.0;

                for (int i = 1; i < pts.Count - 1; i++)
                {
                    XYZ local     = pts[i] - ptA;
                    double along  = local.DotProduct(uAB);
                    XYZ projected = ptA + uAB * along;
                    double dev    = pts[i].DistanceTo(projected);
                    if (dev > maxDev)
                    {
                        maxDev  = dev;
                        bulgePt = pts[i];
                    }
                }

                if (bulgePt == null || maxDev < 1e-6)
                    throw new InvalidOperationException(
                        "FlexPipe appears to be straight — nothing to mirror.");

                // ── 4. Build mirror plane ─────────────────────────────────────
                // bulgeDir = perpendicular-to-AB direction where the curve deviates
                XYZ local2     = bulgePt - ptA;
                double along2  = local2.DotProduct(uAB);
                XYZ projected2 = ptA + uAB * along2;
                XYZ bulgeDir   = (bulgePt - projected2).Normalize();

                // Mirror plane: contains the A-B line, normal = bulgeDir
                // → reflects the bulge to the opposite side of A-B
                Plane mirrorPlane = Plane.CreateByNormalAndOrigin(bulgeDir, ptA);

                // ── 5. Mirror + delete original ───────────────────────────────
                using (Transaction t = new Transaction(doc, "Mirror FlexPipe"))
                {
                    t.Start();

                    FailureHandlingOptions fho = t.GetFailureHandlingOptions();
                    fho.SetFailuresPreprocessor(new WarningSuppressor());
                    t.SetFailureHandlingOptions(fho);

                    // MirrorElement creates a mirrored copy; delete the original
                    ElementTransformUtils.MirrorElement(doc, flexPipe.Id, mirrorPlane);
                    doc.Delete(flexPipe.Id);

                    t.Commit();
                }

                UI?.SetStatus("✔ FlexPipe mirrored successfully.");
                UI?.RestoreWindow();
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                UI?.SetStatus("Mirror canceled by user.");
                UI?.RestoreWindow();
            }
            catch (Exception ex)
            {
                UI?.SetStatus("Mirror error: " + ex.Message);
                UI?.RestoreWindow();
            }
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

        // ── Mirror infrastructure ──────────────────────────────────────────────
        private MirrorFlexPipeHandler _mirrorHandler;
        private ExternalEvent         _mirrorEvent;

        // ── DXF raw deltas (mm) stored on Browse ──────────────────────────────
        private double _dxfDeltaX = double.NaN;
        private double _dxfDeltaY = double.NaN;

        public ConexionFlexPipeWindow(UIApplication uiapp, ElectricalConnectionHandler handler, ExternalEvent exEvent)
        {
            InitializeComponent();
            _uiapp = uiapp;
            _handler = handler;
            _exEvent = exEvent;
            _handler.UI = this;

            // Create dedicated pick handler + event
            _pickHandler = new PickPointHandler { UI = this };
            _pickEvent   = ExternalEvent.Create(_pickHandler);

            // Create mirror handler + event
            _mirrorHandler = new MirrorFlexPipeHandler { UI = this };
            _mirrorEvent   = ExternalEvent.Create(_mirrorHandler);

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
                        System.Windows.Media.Color.FromRgb(0, 130, 60));
                    txtStatus.Text = "Point A selected ✔";
                }
                else
                {
                    _pointB = point;
                    lblElementB.Text = $"Point B: X:{point.X:F3}  Y:{point.Y:F3}  Z:{point.Z:F3}";
                    lblElementB.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(0, 130, 60));
                    txtStatus.Text = "Point B selected ✔";
                }

                UpdateDeltaDisplay(); // ← new

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

            // ── Load writable String parameters from any existing FlexPipe ────
            LoadFlexPipeParameters(doc);
        }

        private void LoadFlexPipeParameters(Document doc)
        {
            var paramNames = new List<string>();

            // Try from an existing FlexPipe instance first
            FlexPipe sample = new FilteredElementCollector(doc)
                .OfClass(typeof(FlexPipe))
                .Cast<FlexPipe>()
                .FirstOrDefault();

            if (sample != null)
            {
                foreach (Parameter p in sample.Parameters)
                {
                    if (!p.IsReadOnly && p.StorageType == StorageType.String
                        && !string.IsNullOrWhiteSpace(p.Definition.Name))
                        paramNames.Add(p.Definition.Name);
                }
            }
            else
            {
                // Fallback: enumerate shared parameters bound to the document
                BindingMap bm = doc.ParameterBindings;
                DefinitionBindingMapIterator it = bm.ForwardIterator();
                it.Reset();
                while (it.MoveNext())
                {
                    Definition def = it.Key;
                    if (def != null && !string.IsNullOrWhiteSpace(def.Name))
                        paramNames.Add(def.Name);
                }
            }

            paramNames = paramNames.Distinct().OrderBy(n => n).ToList();
            cmbParamName.ItemsSource = paramNames;

            // Pre-select "DXF_Source" if present, otherwise first item
            int idx = paramNames.IndexOf("DXF_Source");
            if (idx >= 0)
                cmbParamName.SelectedIndex = idx;
            else if (paramNames.Any())
                cmbParamName.SelectedIndex = 0;
            else
                cmbParamName.Text = "DXF_Source"; // editable fallback
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
            {
                txtDxfPath.Text = ofd.FileName;
                ParseDxfDeltas(ofd.FileName);
                UpdateDeltaDisplay();
            }
        }

        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtDxfPath.Text))
            { SetStatus("Error: Seleccione un archivo DXF."); return; }

            if (_pointA == null || _pointB == null)
            { SetStatus("Error: Debe seleccionar ambos puntos (A y B)."); return; }

            if (cmbFlexPipeType.SelectedItem == null || cmbSystemType.SelectedItem == null)
            { SetStatus("Error: Seleccione los tipos de FlexPipe y Sistema."); return; }

            string paramName = cmbParamName.SelectedItem as string ?? cmbParamName.Text?.Trim();
            if (string.IsNullOrWhiteSpace(paramName))
            { SetStatus("Error: Indique el nombre del parámetro compartido."); return; }

            _handler.DxfFilePath                    = txtDxfPath.Text;
            _handler.PointA                         = _pointA;
            _handler.PointB                         = _pointB;
            _handler.SharedParameterName            = paramName;
            _handler.SelectedFlexPipeTypeKey        = ((ElementType)cmbFlexPipeType.SelectedItem).Name;
            _handler.SelectedElectricalSystemTypeKey = ((ElementType)cmbSystemType.SelectedItem).Name;

            SetStatus("Procesando geometría DXF y generando FlexPipe...");
            _exEvent.Raise();
        }

        private void BtnMirrorFlexPipe_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            _mirrorEvent.Raise();
        }

        // ── Parse only first / last raw XY from DXF (mm, no Revit API needed) ─
        private void ParseDxfDeltas(string filePath)
        {
            _dxfDeltaX = double.NaN;
            _dxfDeltaY = double.NaN;

            try
            {
                DxfFile dxf = DxfFile.Load(filePath);
                double? x0 = null, y0 = null, xN = null, yN = null;

                foreach (DxfEntity entity in dxf.Entities)
                {
                    if (entity is DxfLwPolyline lw && lw.Vertices.Any())
                    {
                        x0 = lw.Vertices.First().X;  y0 = lw.Vertices.First().Y;
                        xN = lw.Vertices.Last().X;   yN = lw.Vertices.Last().Y;
                        break;
                    }
                    if (entity is DxfPolyline poly && poly.Vertices.Any())
                    {
                        x0 = poly.Vertices.First().Location.X;  y0 = poly.Vertices.First().Location.Y;
                        xN = poly.Vertices.Last().Location.X;   yN = poly.Vertices.Last().Location.Y;
                        break;
                    }
                    if (entity is DxfLine line)
                    {
                        if (x0 == null) { x0 = line.P1.X; y0 = line.P1.Y; }
                        xN = line.P2.X; yN = line.P2.Y;
                    }
                    if (entity is DxfSpline spline && spline.ControlPoints.Any())
                    {
                        x0 = spline.ControlPoints.First().Point.X;  y0 = spline.ControlPoints.First().Point.Y;
                        xN = spline.ControlPoints.Last().Point.X;   yN = spline.ControlPoints.Last().Point.Y;
                        break;
                    }
                }

                if (x0 != null && xN != null)
                {
                    _dxfDeltaX = Math.Abs(xN.Value - x0.Value);
                    _dxfDeltaY = (yN.Value - y0.Value);
                }
            }
            catch { /* silent — display will show N/A */ }
        }

        // ── Compute and display both delta rows ───────────────────────────────
        private void UpdateDeltaDisplay()
        {
            // DXF row
            string dxfRow = double.IsNaN(_dxfDeltaX)
                ? "DXF:  —"
                : $"DXF:  ΔX = {_dxfDeltaX:F1} mm   ΔY = {_dxfDeltaY:F1} mm";

            // Model row — project B-A into horizontal / vertical components
            string modelRow;
            if (_pointA != null && _pointB != null)
            {
                double ftToMm = UnitUtils.ConvertFromInternalUnits(1.0, UnitTypeId.Millimeters);

                double horizontalMm = Math.Sqrt(
                    Math.Pow((_pointB.X - _pointA.X) * ftToMm, 2) +
                    Math.Pow((_pointB.Y - _pointA.Y) * ftToMm, 2));
                double verticalMm = (_pointB.Z - _pointA.Z) * ftToMm;

                modelRow = $"MODEL: ΔX' = {horizontalMm:F1} mm   ΔY' = {verticalMm:F1} mm";
            }
            else
            {
                modelRow = "MODEL: —";
            }

            lblDeltaInfo.Text = dxfRow + "\n" + modelRow;
        }
    }
}