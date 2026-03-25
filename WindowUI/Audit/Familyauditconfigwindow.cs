using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

using Color = System.Windows.Media.Color;
using Button = System.Windows.Controls.Button;
using CheckBox = System.Windows.Controls.CheckBox;
using ComboBox = System.Windows.Controls.ComboBox;
using TextBox = System.Windows.Controls.TextBox;

namespace HMVTools
{
    // ═══════════════════════════════════════════════════════════════
    //  Plain data class — config chosen by the user, passed to command
    // ═══════════════════════════════════════════════════════════════
    public class FamilyAuditConfig
    {
        public string AdcRootPath { get; set; } = "";
        public AuditScope Scope { get; set; } = AuditScope.AllFamilies;
        public string CategoryFilter { get; set; } = "";   // only when Scope == ByCategory
        public bool InjectParameter { get; set; } = false;
        public bool EnableProactiveTracking { get; set; } = true;
    }

    public enum AuditScope
    {
        AllFamilies,
        SelectedFamilies,
        ByCategory
    }

    // ═══════════════════════════════════════════════════════════════
    //  PRE-AUDIT CONFIG WINDOW
    // ═══════════════════════════════════════════════════════════════
    public class FamilyAuditConfigWindow : Window
    {
        public FamilyAuditConfig ResultConfig { get; private set; }

        private TextBox _txtAdcRoot;
        private ComboBox _cmbScope;
        private ComboBox _cmbCategory;
        private CheckBox _chkInject;
        private CheckBox _chkProactive;

        private readonly List<string> _categories;
        private readonly int _selectedCount;
        private readonly bool _hasSelection;

        // ── Persisted across sessions (static, survives command re-runs) ──
        private static string _lastAdcRoot = null;
        private static bool _lastInject = false;
        private static bool _lastProactive = true;

        static readonly Color COL_BG = Color.FromRgb(245, 245, 248);
        static readonly Color COL_TEXT = Color.FromRgb(30, 30, 30);
        static readonly Color COL_SUB = Color.FromRgb(120, 120, 130);
        static readonly Color COL_ACCENT = Color.FromRgb(0, 120, 212);
        static readonly Color COL_BORDER = Color.FromRgb(200, 200, 210);
        static readonly Color COL_WARN = Color.FromRgb(220, 53, 69);

        /// <param name="categories">Distinct family category names in the document.</param>
        /// <param name="selectedCount">Number of pre-selected family instances (0 = none).</param>
        public FamilyAuditConfigWindow(List<string> categories, int selectedCount)
        {
            _categories = categories ?? new List<string>();
            _selectedCount = selectedCount;
            _hasSelection = selectedCount > 0;

            Title = "HMV Tools - Family Audit";
            Width = 520; Height = 480;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(COL_BG);

            BuildUI();
        }

        void BuildUI()
        {
            StackPanel root = new StackPanel { Margin = new Thickness(24) };
            Content = root;

            // ── Header ────────────────────────────────────────────
            root.Children.Add(new TextBlock
            {
                Text = "Family Audit — Configuration",
                FontSize = 20,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(COL_TEXT),
                Margin = new Thickness(0, 0, 0, 4)
            });
            root.Children.Add(new TextBlock
            {
                Text = "Configure the audit scope and ADC workspace path.",
                FontSize = 12,
                Foreground = new SolidColorBrush(COL_SUB),
                Margin = new Thickness(0, 0, 0, 18)
            });

            // ── ADC Root Path ─────────────────────────────────────
            root.Children.Add(SectionLabel("ADC Workspace Root"));
            root.Children.Add(new TextBlock
            {
                Text = "Folder searched for .rfa source files (Tier 2 fallback).",
                FontSize = 10,
                Foreground = new SolidColorBrush(COL_SUB),
                Margin = new Thickness(0, 0, 0, 4)
            });

            StackPanel pathRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 14)
            };
            Border pathBorder = InputBorder();
            _txtAdcRoot = new TextBox
            {
                Width = 380,
                Height = 28,
                FontSize = 11,
                VerticalContentAlignment = VerticalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(6, 0, 6, 0),
                Text = _lastAdcRoot ?? DefaultAdcRoot()
            };
            WireFocus(_txtAdcRoot, pathBorder);
            pathBorder.Child = _txtAdcRoot;
            pathRow.Children.Add(pathBorder);

            Button btnBrowse = MkBtn("...", COL_ACCENT, Colors.White, 36);
            btnBrowse.Height = 28; btnBrowse.Margin = new Thickness(6, 0, 0, 0);
            btnBrowse.ToolTip = "Browse for ADC folder";
            btnBrowse.Click += (s, e) =>
            {
                var dlg = new System.Windows.Forms.FolderBrowserDialog
                {
                    Description = "Select ADC workspace root folder",
                    ShowNewFolderButton = false
                };
                if (!string.IsNullOrEmpty(_txtAdcRoot.Text)
                    && Directory.Exists(_txtAdcRoot.Text))
                    dlg.SelectedPath = _txtAdcRoot.Text;

                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    _txtAdcRoot.Text = dlg.SelectedPath;
            };
            pathRow.Children.Add(btnBrowse);
            root.Children.Add(pathRow);

            // ── Scope ─────────────────────────────────────────────
            root.Children.Add(SectionLabel("Audit Scope"));

            _cmbScope = new ComboBox
            {
                Width = 430,
                Height = 30,
                FontSize = 12,
                VerticalContentAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(0, 0, 0, 6)
            };
            _cmbScope.Items.Add("All editable families");
            _cmbScope.Items.Add($"Selected families only ({_selectedCount} selected)");
            _cmbScope.Items.Add("Filter by category");
            _cmbScope.SelectedIndex = 0;
            _cmbScope.SelectionChanged += (s, e) => OnScopeChanged();

            // Disable "Selected" option if nothing is pre-selected
            if (!_hasSelection)
            {
                var item = _cmbScope.Items[1] as string;
                // We'll just show 0 selected and validate on apply
            }

            root.Children.Add(_cmbScope);

            // Category dropdown (visible only when scope = ByCategory)
            _cmbCategory = new ComboBox
            {
                Width = 430,
                Height = 30,
                FontSize = 12,
                VerticalContentAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(0, 0, 0, 14),
                Visibility = Visibility.Collapsed
            };
            foreach (string cat in _categories.OrderBy(c => c))
                _cmbCategory.Items.Add(cat);
            if (_cmbCategory.Items.Count > 0)
                _cmbCategory.SelectedIndex = 0;
            root.Children.Add(_cmbCategory);

            // ── Options ───────────────────────────────────────────
            root.Children.Add(SectionLabel("Options"));

            _chkInject = new CheckBox
            {
                Content = "Inject ADC_Library_Path type parameter into families",
                FontSize = 11.5,
                Foreground = new SolidColorBrush(COL_TEXT),
                IsChecked = _lastInject,
                Margin = new Thickness(0, 0, 0, 4)
            };
            root.Children.Add(_chkInject);

            root.Children.Add(new TextBlock
            {
                Text = "⚠ Heavyweight: opens each family in the editor and reloads it (~1-3 s per family).",
                FontSize = 9.5,
                Foreground = new SolidColorBrush(COL_WARN),
                Margin = new Thickness(18, 0, 0, 10),
                TextWrapping = TextWrapping.Wrap
            });

            _chkProactive = new CheckBox
            {
                Content = "Enable proactive tracking on family load events",
                FontSize = 11.5,
                Foreground = new SolidColorBrush(COL_TEXT),
                IsChecked = _lastProactive,
                Margin = new Thickness(0, 0, 0, 14)
            };
            root.Children.Add(_chkProactive);

            // ── Buttons ───────────────────────────────────────────
            StackPanel btnRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 6, 0, 0)
            };
            Button btnCancel = MkBtn("Cancelar", Color.FromRgb(240, 240, 243), COL_TEXT, 100);
            btnCancel.Click += (s, e) => { DialogResult = false; Close(); };
            btnRow.Children.Add(btnCancel);

            Button btnGo = MkBtn("Run Audit →", COL_ACCENT, Colors.White, 140);
            btnGo.FontWeight = FontWeights.SemiBold;
            btnGo.Click += DoAccept;
            btnRow.Children.Add(btnGo);
            root.Children.Add(btnRow);
        }

        // ══════════════════════════════════════════════════════════
        //  LOGIC
        // ══════════════════════════════════════════════════════════

        void OnScopeChanged()
        {
            _cmbCategory.Visibility =
                _cmbScope.SelectedIndex == 2
                    ? Visibility.Visible
                    : Visibility.Collapsed;
        }

        void DoAccept(object sender, RoutedEventArgs e)
        {
            string adcRoot = _txtAdcRoot.Text.Trim();

            // Validate ADC path
            if (string.IsNullOrEmpty(adcRoot) || !Directory.Exists(adcRoot))
            {
                MessageBox.Show(
                    "The ADC workspace path does not exist.\n\n"
                    + "Please enter a valid folder path or click \"...\" to browse.",
                    "HMV - Family Audit",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                _txtAdcRoot.Focus();
                return;
            }

            // Validate selected scope
            AuditScope scope;
            switch (_cmbScope.SelectedIndex)
            {
                case 1:
                    if (!_hasSelection)
                    {
                        MessageBox.Show(
                            "No family instances are selected in the canvas.\n\n"
                            + "Pre-select families before running the command, "
                            + "or choose a different scope.",
                            "HMV - Family Audit",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    scope = AuditScope.SelectedFamilies;
                    break;
                case 2:
                    if (_cmbCategory.SelectedItem == null)
                    {
                        MessageBox.Show(
                            "Please select a category to filter by.",
                            "HMV - Family Audit",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    scope = AuditScope.ByCategory;
                    break;
                default:
                    scope = AuditScope.AllFamilies;
                    break;
            }

            // Persist for next run
            _lastAdcRoot = adcRoot;
            _lastInject = _chkInject.IsChecked == true;
            _lastProactive = _chkProactive.IsChecked == true;

            ResultConfig = new FamilyAuditConfig
            {
                AdcRootPath = adcRoot,
                Scope = scope,
                CategoryFilter = _cmbCategory.SelectedItem as string ?? "",
                InjectParameter = _chkInject.IsChecked == true,
                EnableProactiveTracking = _chkProactive.IsChecked == true
            };

            DialogResult = true;
            Close();
        }

        // ══════════════════════════════════════════════════════════
        //  HELPERS
        // ══════════════════════════════════════════════════════════

        static string DefaultAdcRoot()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "DC", "ACCDocs");
        }

        TextBlock SectionLabel(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(COL_ACCENT),
                Margin = new Thickness(0, 0, 0, 4)
            };
        }

        Border InputBorder()
        {
            return new Border
            {
                CornerRadius = new CornerRadius(5),
                BorderBrush = new SolidColorBrush(COL_BORDER),
                BorderThickness = new Thickness(1),
                Background = Brushes.White
            };
        }

        void WireFocus(TextBox t, Border b)
        {
            t.GotFocus += (s, e) =>
                b.BorderBrush = new SolidColorBrush(COL_ACCENT);
            t.LostFocus += (s, e) =>
                b.BorderBrush = new SolidColorBrush(COL_BORDER);
        }

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
            cp.SetValue(ContentPresenter.HorizontalAlignmentProperty,
                HorizontalAlignment.Center);
            cp.SetValue(ContentPresenter.VerticalAlignmentProperty,
                VerticalAlignment.Center);
            bd.AppendChild(cp);
            tp.VisualTree = bd;
            b.Template = tp;
            return b;
        }
    }
}