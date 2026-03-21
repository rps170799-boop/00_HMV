using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

using Color = System.Windows.Media.Color;
using TextBox = System.Windows.Controls.TextBox;
using ComboBox = System.Windows.Controls.ComboBox;
using CheckBox = System.Windows.Controls.CheckBox;
using Button = System.Windows.Controls.Button;

namespace HMVTools
{
    public class Dwg3DToShapeWindow : Window
    {
        // ── Results read by the command ──
        public double ThresholdCm3 { get; private set; } = 1.0;
        public int SelectedCategoryBic { get; private set; }
        public Autodesk.Revit.DB.ElementId SelectedImportId { get; private set; }
        public bool DeleteOriginal { get; private set; }
        public string ShapeName { get; private set; } = "DWG_DirectShape";

        private readonly List<Dwg3DImportItem> _imports;
        private readonly List<Dwg3DCategoryChoice> _categories;

        private ComboBox _cmbImport, _cmbCategory;
        private TextBox _txtThreshold, _txtName;
        private CheckBox _chkDelete;

        static readonly Color CA = Color.FromRgb(0, 120, 212);
        static readonly Color CT = Color.FromRgb(30, 30, 30);
        static readonly Color CS = Color.FromRgb(120, 120, 130);
        static readonly Color CB = Color.FromRgb(200, 200, 210);
        static readonly Color CBG = Color.FromRgb(245, 245, 248);

        public Dwg3DToShapeWindow(
            List<Dwg3DImportItem> imports,
            List<Dwg3DCategoryChoice> categories)
        {
            _imports = imports;
            _categories = categories;
            BuildUI();
        }

        void BuildUI()
        {
            Title = "HMV Tools - 3D DWG to DirectShape";
            Width = 480; Height = 520;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(CBG);

            StackPanel root = new StackPanel { Margin = new Thickness(24) };
            Content = root;

            // ── Header ──
            root.Children.Add(new TextBlock
            {
                Text = "3D DWG → DirectShape",
                FontSize = 20,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(CT),
                Margin = new Thickness(0, 0, 0, 4)
            });
            root.Children.Add(new TextBlock
            {
                Text = "Extracts solids and meshes from an imported DWG\nand creates a native DirectShape element.",
                FontSize = 11,
                Foreground = new SolidColorBrush(CS),
                Margin = new Thickness(0, 0, 0, 18),
                TextWrapping = TextWrapping.Wrap
            });

            // ── DWG Import selector ──
            root.Children.Add(Lbl("DWG Import"));
            _cmbImport = new ComboBox
            {
                Height = 32,
                FontSize = 12,
                VerticalContentAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 14)
            };
            foreach (var item in _imports) _cmbImport.Items.Add(item);
            _cmbImport.SelectedIndex = 0;
            root.Children.Add(_cmbImport);

            // ── Category selector ──
            root.Children.Add(Lbl("Target Category"));
            _cmbCategory = new ComboBox
            {
                Height = 32,
                FontSize = 12,
                VerticalContentAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 14)
            };
            foreach (var cat in _categories) _cmbCategory.Items.Add(cat);
            _cmbCategory.SelectedIndex = 0;
            root.Children.Add(_cmbCategory);

            // ── Threshold + Name row ──
            Grid row = new Grid { Margin = new Thickness(0, 0, 0, 14) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(16) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            root.Children.Add(row);

            // Threshold
            StackPanel thSp = new StackPanel();
            thSp.Children.Add(Lbl("Volume Threshold (cm³)"));
            _txtThreshold = new TextBox
            {
                Height = 30,
                FontSize = 12,
                Text = "1.0",
                VerticalContentAlignment = VerticalAlignment.Center,
                Padding = new Thickness(6, 0, 6, 0),
                ToolTip = "Solids with volume below this value are skipped\n(bolts, nuts, small details)"
            };
            thSp.Children.Add(_txtThreshold);
            Grid.SetColumn(thSp, 0);
            row.Children.Add(thSp);

            // Name
            StackPanel nmSp = new StackPanel();
            nmSp.Children.Add(Lbl("DirectShape Name"));
            _txtName = new TextBox
            {
                Height = 30,
                FontSize = 12,
                Text = "DWG_DirectShape",
                VerticalContentAlignment = VerticalAlignment.Center,
                Padding = new Thickness(6, 0, 6, 0)
            };
            nmSp.Children.Add(_txtName);
            Grid.SetColumn(nmSp, 2);
            row.Children.Add(nmSp);

            // ── Delete original checkbox ──
            _chkDelete = new CheckBox
            {
                Content = "Delete original DWG import after conversion",
                FontSize = 11,
                IsChecked = false,
                Foreground = new SolidColorBrush(CT),
                Margin = new Thickness(0, 0, 0, 24)
            };
            root.Children.Add(_chkDelete);

            // ── Info box ──
            Border infoBox = new Border
            {
                CornerRadius = new CornerRadius(6),
                Background = new SolidColorBrush(Color.FromRgb(232, 243, 255)),
                Padding = new Thickness(12, 8, 12, 8),
                Margin = new Thickness(0, 0, 0, 20)
            };
            infoBox.Child = new TextBlock
            {
                Text = "Tip: The threshold filters out small solids (bolts, washers) " +
                       "by volume. Meshes are always converted regardless of size. " +
                       "If your DWG has only meshes, set threshold to 0.",
                FontSize = 10.5,
                Foreground = new SolidColorBrush(CA),
                TextWrapping = TextWrapping.Wrap
            };
            root.Children.Add(infoBox);

            // ── Buttons ──
            StackPanel btnRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            Button btnCancel = MkBtn("Cancel", Color.FromRgb(240, 240, 243), CT, 90);
            btnCancel.Click += (s, e) => { DialogResult = false; Close(); };
            btnRow.Children.Add(btnCancel);

            Button btnRun = MkBtn("Convert →", CA, Colors.White, 130);
            btnRun.FontWeight = FontWeights.SemiBold;
            btnRun.Click += DoRun;
            btnRow.Children.Add(btnRun);

            root.Children.Add(btnRow);
        }

        void DoRun(object sender, RoutedEventArgs e)
        {
            if (!double.TryParse(_txtThreshold.Text, out double th) || th < 0)
            {
                MessageBox.Show("Enter a valid threshold (≥ 0).", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_cmbImport.SelectedItem == null)
            {
                MessageBox.Show("Select a DWG import.", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            ThresholdCm3 = th;
            SelectedCategoryBic = ((Dwg3DCategoryChoice)_cmbCategory.SelectedItem).BicInt;
            SelectedImportId = ((Dwg3DImportItem)_cmbImport.SelectedItem).Id;
            DeleteOriginal = _chkDelete.IsChecked == true;
            ShapeName = string.IsNullOrWhiteSpace(_txtName.Text)
                                    ? "DWG_DirectShape" : _txtName.Text.Trim();

            DialogResult = true;
            Close();
        }

        // ── UI helpers ──
        TextBlock Lbl(string text) => new TextBlock
        {
            Text = text,
            FontSize = 12,
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(CA),
            Margin = new Thickness(0, 0, 0, 4)
        };

        Button MkBtn(string text, Color bg, Color fg, double w)
        {
            Button b = new Button
            {
                Content = text,
                Width = w,
                Height = 36,
                FontSize = 13,
                Margin = new Thickness(0, 0, 8, 0),
                Cursor = Cursors.Hand,
                Foreground = new SolidColorBrush(fg),
                Background = new SolidColorBrush(bg),
                BorderThickness = new Thickness(0)
            };
            var tp = new ControlTemplate(typeof(Button));
            var bd = new FrameworkElementFactory(typeof(Border));
            bd.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
            bd.SetValue(Border.BackgroundProperty, new SolidColorBrush(bg));
            bd.SetValue(Border.PaddingProperty, new Thickness(14, 6, 14, 6));
            var cp = new FrameworkElementFactory(typeof(ContentPresenter));
            cp.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            cp.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            bd.AppendChild(cp); tp.VisualTree = bd; b.Template = tp;
            return b;
        }
    }
}