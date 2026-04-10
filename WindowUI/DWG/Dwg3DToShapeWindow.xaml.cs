using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;

namespace HMVTools
{
    // ── Output mode enum used by the command ────────────────────────
    public enum Dwg3DOutputMode
    {
        DirectShapeInProject,
        NewFamilyFromDwg,
        MigrateDsToFamily
    }

    public class Dwg3DShapeItem
    {
        public ElementId Id { get; set; }
        public string Name { get; set; }
        public override string ToString() => Name;
    }

    public partial class Dwg3DToShapeWindow : Window
    {
        // ── Public properties read by the Revit command ─────────────
        public double ThresholdCm3 { get; private set; } = 1.0;
        public double MeshBBoxMm { get; private set; } = 50.0;
        public int DecimateFactor { get; private set; } = 1;
        public int SelectedCategoryBic { get; private set; }
        public ElementId SelectedImportId { get; private set; }
        public ElementId SelectedSourceDsId { get; private set; }
        public bool DeleteOriginal { get; private set; }
        public bool CreateReferencePlanes { get; private set; } = true;
        public Dwg3DOutputMode Mode { get; private set; } = Dwg3DOutputMode.DirectShapeInProject;
        public string ShapeName { get; private set; } = "DWG_DirectShape";

        // ── Backing lists ────────────────────────────────────────────
        private readonly List<Dwg3DImportItem> _imports;
        private readonly List<Dwg3DCategoryChoice> _categories;
        private readonly List<Dwg3DShapeItem> _directShapes;

        public Dwg3DToShapeWindow(
            List<Dwg3DImportItem> imports,
            List<Dwg3DCategoryChoice> categories,
            List<Dwg3DShapeItem> directShapes)
        {
            InitializeComponent();

            _imports = imports ?? new List<Dwg3DImportItem>();
            _categories = categories ?? new List<Dwg3DCategoryChoice>();
            _directShapes = directShapes ?? new List<Dwg3DShapeItem>();

            // ── Populate Mode combo ──
            cmbMode.Items.Add("DirectShape in Project");
            cmbMode.Items.Add("New Family from DWG");
            cmbMode.Items.Add("Migrate DirectShape → Family");
            cmbMode.SelectedIndex = 0;

            // ── Populate Categories (bind objects, not strings) ──
            foreach (var cat in _categories) cmbCategory.Items.Add(cat);
            if (cmbCategory.Items.Count > 0) cmbCategory.SelectedIndex = 0;

            // ── Populate DWG Imports (bind objects) ──
            foreach (var imp in _imports) cmbImports.Items.Add(imp);
            if (cmbImports.Items.Count > 0) cmbImports.SelectedIndex = 0;

            // ── Populate DirectShapes (bind objects) ──
            foreach (var ds in _directShapes) cmbSourceDs.Items.Add(ds);
            if (cmbSourceDs.Items.Count > 0) cmbSourceDs.SelectedIndex = 0;

            // ── Pre-fill shape name from first DWG's base filename ──
            string initialName = "DWG_DirectShape";
            if (_imports.Count > 0 && !string.IsNullOrWhiteSpace(_imports[0].BaseFileName))
                initialName = _imports[0].BaseFileName;
            txtShapeName.Text = initialName;

            UpdateModeVisibility();
        }

        // ── Visibility / label updates based on current mode ─────────
        private void UpdateModeVisibility()
        {
            bool isModeC = cmbMode.SelectedIndex == 2;
            bool isFamilyMode = cmbMode.SelectedIndex == 1 || isModeC;

            pnlImport.Visibility = isModeC ? System.Windows.Visibility.Collapsed : System.Windows.Visibility.Visible;
            pnlSourceDs.Visibility = isModeC ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            chkRefPlanes.Visibility = isFamilyMode ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

            chkDeleteOriginal.Content = isModeC
                ? "Delete source DirectShape after migration"
                : "Delete original DWG import after conversion";
        }

        // ── Handlers ─────────────────────────────────────────────────
        private void CmbMode_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (IsLoaded || pnlImport != null) UpdateModeVisibility();
        }

        private void CmbImports_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbImports.SelectedItem is Dwg3DImportItem item
                && !string.IsNullOrWhiteSpace(item.BaseFileName)
                && txtShapeName != null)
            {
                txtShapeName.Text = item.BaseFileName;
            }
        }

        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void NumericOnly(object sender, TextCompositionEventArgs e)
        {
            TextBox tb = sender as TextBox;
            foreach (char c in e.Text)
            {
                if (!char.IsDigit(c) && c != '.') { e.Handled = true; return; }
                if (c == '.' && tb != null && tb.Text.Contains(".")) { e.Handled = true; return; }
            }
        }

        // ── Run / validate ──────────────────────────────────────────
        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            if (!double.TryParse(txtThreshold.Text, out double th) || th < 0)
            {
                MessageBox.Show("Enter a valid volume threshold (≥ 0).", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!double.TryParse(txtMaxEdge.Text, out double bb) || bb < 0)
            {
                MessageBox.Show("Enter a valid mesh BBox threshold (≥ 0).", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!int.TryParse(txtDecimate.Text, out int dec) || dec < 1)
            {
                MessageBox.Show("Enter a valid decimate factor (≥ 1).", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Resolve mode
            Mode = cmbMode.SelectedIndex == 1 ? Dwg3DOutputMode.NewFamilyFromDwg
                 : cmbMode.SelectedIndex == 2 ? Dwg3DOutputMode.MigrateDsToFamily
                 : Dwg3DOutputMode.DirectShapeInProject;

            // Mode-specific source selection
            if (Mode == Dwg3DOutputMode.MigrateDsToFamily)
            {
                if (!(cmbSourceDs.SelectedItem is Dwg3DShapeItem dsItem))
                {
                    MessageBox.Show("Select a source DirectShape.", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                SelectedSourceDsId = dsItem.Id;
            }
            else
            {
                if (!(cmbImports.SelectedItem is Dwg3DImportItem impItem))
                {
                    MessageBox.Show("Select a DWG import.", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                SelectedImportId = impItem.Id;
            }

            // Category
            if (cmbCategory.SelectedItem is Dwg3DCategoryChoice catItem)
                SelectedCategoryBic = catItem.BicInt;
            else
                SelectedCategoryBic = (int)BuiltInCategory.OST_GenericModel;

            // Scalars
            ThresholdCm3 = th;
            MeshBBoxMm = bb;
            DecimateFactor = dec;

            // Name
            ShapeName = string.IsNullOrWhiteSpace(txtShapeName.Text)
                ? "DWG_DirectShape"
                : txtShapeName.Text.Trim();

            // Flags
            DeleteOriginal = chkDeleteOriginal.IsChecked == true;
            CreateReferencePlanes = chkRefPlanes.IsChecked == true;

            DialogResult = true;
            Close();
        }
    }
}