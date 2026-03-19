using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Autodesk.Revit.DB;

using Color = System.Windows.Media.Color;
using Grid = System.Windows.Controls.Grid;

namespace HMVTools
{
    public class MultiParamEditorWindow : Window
    {
        // Result: familyKey → list of changes
        public Dictionary<string, List<ParamChange>> ResultChanges
        { get; private set; }

        // Data
        private readonly Dictionary<string, List<ParamInfo>>
            _familyParams;
        private readonly Dictionary<string, List<FamilyInstance>>
            _familyGroups;
        private readonly bool _commonMode;
        private readonly List<string> _familyKeys;

        // Per-family persistent rows: familyKey → list of row data
        private readonly Dictionary<string, List<ParamRowData>>
            _familyRows;

        private string _activeFamilyKey;

        // UI
        private StackPanel _rowsPanel;
        private StackPanel _radioPanel;
        private TextBlock _headerLabel;

        // Colors
        private static readonly Color BluePrimary =
            Color.FromRgb(0, 120, 212);
        private static readonly Color GreenBtn =
            Color.FromRgb(40, 167, 69);
        private static readonly Color GrayBg =
            Color.FromRgb(240, 240, 243);
        private static readonly Color DarkText =
            Color.FromRgb(30, 30, 30);
        private static readonly Color MutedText =
            Color.FromRgb(120, 120, 130);
        private static readonly Color PlaceholderGray =
            Color.FromRgb(160, 160, 160);
        private static readonly Color BorderClr =
            Color.FromRgb(200, 200, 210);
        private static readonly Color WindowBg =
            Color.FromRgb(245, 245, 248);
        private static readonly Color RedBtn =
            Color.FromRgb(220, 53, 69);

        private class ParamRowData
        {
            public string SelectedParam;
            public string ValueText;
        }

        public MultiParamEditorWindow(
            Dictionary<string, List<ParamInfo>> familyParams,
            Dictionary<string, List<FamilyInstance>> familyGroups,
            bool commonMode)
        {
            _familyParams = familyParams;
            _familyGroups = familyGroups;
            _commonMode = commonMode;
            _familyKeys = familyParams.Keys.ToList();

            _familyRows =
                new Dictionary<string, List<ParamRowData>>();
            foreach (string key in _familyKeys)
                _familyRows[key] = new List<ParamRowData>();

            _activeFamilyKey = _familyKeys[0];

            BuildUI();
        }

        private void BuildUI()
        {
            Title = "HMV Tools – Multi Instance-Parameter Editor";
            Width = 600;
            MinWidth = 500;
            Height = 500;
            MinHeight = 350;
            WindowStartupLocation =
                WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.CanResize;
            Background = new SolidColorBrush(WindowBg);

            var root = new Grid
            {
                Margin = new Thickness(24)
            };
            root.RowDefinitions.Add(new RowDefinition
            { Height = GridLength.Auto }); // 0 title
            root.RowDefinitions.Add(new RowDefinition
            { Height = GridLength.Auto }); // 1 radio buttons
            root.RowDefinitions.Add(new RowDefinition
            { Height = GridLength.Auto }); // 2 header
            root.RowDefinitions.Add(new RowDefinition
            { Height = new GridLength(1, GridUnitType.Star) }); // 3 rows
            root.RowDefinitions.Add(new RowDefinition
            { Height = GridLength.Auto }); // 4 add button
            root.RowDefinitions.Add(new RowDefinition
            { Height = GridLength.Auto }); // 5 action buttons
            Content = root;

            // ── Row 0: Title ──
            var title = new TextBlock
            {
                Text = "Multi Instance-Parameter Editor",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 4)
            };
            Grid.SetRow(title, 0);
            root.Children.Add(title);

            // ── Row 1: Radio buttons (each-family mode only) ──
            _radioPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 10)
            };

            if (!_commonMode && _familyKeys.Count > 1)
            {
                foreach (string key in _familyKeys)
                {
                    int count = _familyGroups.ContainsKey(key)
                        ? _familyGroups[key].Count : 0;
                    var rb = new RadioButton
                    {
                        Content = key + " (" + count + ")",
                        FontSize = 12,
                        Foreground = new SolidColorBrush(DarkText),
                        Margin = new Thickness(0, 0, 16, 0),
                        Tag = key,
                        GroupName = "FamilySelect",
                        IsChecked = key == _activeFamilyKey
                    };
                    rb.Checked += OnFamilyRadioChecked;
                    _radioPanel.Children.Add(rb);
                }
            }
            else
            {
                int totalCount = _familyGroups.Values
                    .Sum(l => l.Count);
                string label = _commonMode
                    ? "Common parameters – " + totalCount
                      + " instance(s) across "
                      + _familyGroups.Count + " families"
                    : _familyKeys[0] + " – "
                      + totalCount + " instance(s)";
                _radioPanel.Children.Add(new TextBlock
                {
                    Text = label,
                    FontSize = 12,
                    Foreground = new SolidColorBrush(MutedText)
                });
            }
            Grid.SetRow(_radioPanel, 1);
            root.Children.Add(_radioPanel);

            // ── Row 2: Column header ──
            _headerLabel = new TextBlock
            {
                Text = "Parameter                              "
                     + "New Value",
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(0, 0, 0, 4)
            };
            Grid.SetRow(_headerLabel, 2);
            root.Children.Add(_headerLabel);

            // ── Row 3: Scrollable parameter rows ──
            var scroller = new ScrollViewer
            {
                VerticalScrollBarVisibility =
                    ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility =
                    ScrollBarVisibility.Disabled,
                Margin = new Thickness(0, 0, 0, 8)
            };
            _rowsPanel = new StackPanel();
            scroller.Content = _rowsPanel;
            Grid.SetRow(scroller, 3);
            root.Children.Add(scroller);

            // ── Row 4: Add button ──
            var addBtn = MakeButton("＋  Add Parameter",
                GrayBg, DarkText, 160);
            addBtn.HorizontalAlignment =
                HorizontalAlignment.Left;
            addBtn.Margin = new Thickness(0, 0, 0, 12);
            addBtn.Click += (s, e) => AddRow(null, "");
            Grid.SetRow(addBtn, 4);
            root.Children.Add(addBtn);

            // ── Row 5: Action buttons ──
            var btnRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var cancelBtn = MakeButton("Cancel",
                GrayBg, Color.FromRgb(60, 60, 60), 90);
            cancelBtn.Click += (s, e) =>
            {
                DialogResult = false;
                Close();
            };
            btnRow.Children.Add(cancelBtn);

            var applyBtn = MakeButton("Apply Changes",
                GreenBtn, Colors.White, 160);
            applyBtn.FontWeight = FontWeights.SemiBold;
            applyBtn.Click += OnApply;
            btnRow.Children.Add(applyBtn);

            Grid.SetRow(btnRow, 5);
            root.Children.Add(btnRow);

            // Start with one empty row
            AddRow(null, "");
        }

        // ═══════════════════════════════════════════════════════
        // Family switching
        // ═══════════════════════════════════════════════════════

        private void OnFamilyRadioChecked(
            object sender, RoutedEventArgs e)
        {
            var rb = sender as RadioButton;
            if (rb == null) return;
            string newKey = rb.Tag as string;
            if (newKey == _activeFamilyKey) return;

            SaveCurrentRows();
            _activeFamilyKey = newKey;
            LoadRowsForFamily();
        }

        private void SaveCurrentRows()
        {
            var saved = new List<ParamRowData>();
            foreach (UIElement child in _rowsPanel.Children)
            {
                var rowGrid = child as Grid;
                if (rowGrid == null) continue;

                var combo = FindChild<ComboBox>(rowGrid);
                var textBox = FindChild<TextBox>(rowGrid);
                if (combo == null) continue;

                saved.Add(new ParamRowData
                {
                    SelectedParam =
                        combo.SelectedItem as string,
                    ValueText = textBox != null
                        ? textBox.Text : ""
                });
            }
            _familyRows[_activeFamilyKey] = saved;
        }

        private void LoadRowsForFamily()
        {
            _rowsPanel.Children.Clear();
            var rows = _familyRows[_activeFamilyKey];
            if (rows.Count == 0)
            {
                AddRow(null, "");
            }
            else
            {
                foreach (var rd in rows)
                    AddRow(rd.SelectedParam, rd.ValueText);
            }
        }

        // ═══════════════════════════════════════════════════════
        // Row management
        // ═══════════════════════════════════════════════════════

        private void AddRow(string selectedParam,
            string valueText)
        {
            var rowGrid = new Grid
            {
                Margin = new Thickness(0, 0, 0, 6)
            };
            rowGrid.ColumnDefinitions.Add(
                new ColumnDefinition
                {
                    Width = new GridLength(1, GridUnitType.Star)
                });
            rowGrid.ColumnDefinitions.Add(
                new ColumnDefinition { Width = new GridLength(8) });
            rowGrid.ColumnDefinitions.Add(
                new ColumnDefinition
                {
                    Width = new GridLength(1, GridUnitType.Star)
                });
            rowGrid.ColumnDefinitions.Add(
                new ColumnDefinition
                {
                    Width = new GridLength(50)
                });
            rowGrid.ColumnDefinitions.Add(
                new ColumnDefinition
                {
                    Width = new GridLength(36)
                });

            // Combo
            var comboBorder = new Border
            {
                CornerRadius = new CornerRadius(6),
                BorderBrush = new SolidColorBrush(BorderClr),
                BorderThickness = new Thickness(1),
                Background = Brushes.White
            };
            var combo = new ComboBox
            {
                Height = 32,
                FontSize = 12,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(6, 0, 6, 0),
                VerticalContentAlignment =
                    VerticalAlignment.Center
            };

            List<ParamInfo> paramList =
                _familyParams[_activeFamilyKey];
            foreach (ParamInfo pi in paramList)
                combo.Items.Add(pi.Name);

            if (selectedParam != null
                && combo.Items.Contains(selectedParam))
                combo.SelectedItem = selectedParam;

            comboBorder.Child = combo;
            Grid.SetColumn(comboBorder, 0);
            rowGrid.Children.Add(comboBorder);

            // Value textbox
            var valBorder = new Border
            {
                CornerRadius = new CornerRadius(6),
                BorderBrush = new SolidColorBrush(BorderClr),
                BorderThickness = new Thickness(1),
                Background = Brushes.White
            };
            var valBox = new TextBox
            {
                Height = 32,
                FontSize = 12,
                VerticalContentAlignment =
                    VerticalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(8, 0, 8, 0),
                Text = valueText ?? ""
            };

            valBorder.Child = valBox;
            Grid.SetColumn(valBorder, 2);
            rowGrid.Children.Add(valBorder);

            // Unit label
            var unitLabel = new TextBlock
            {
                Text = "",
                FontSize = 10,
                Foreground = new SolidColorBrush(MutedText),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment =
                    HorizontalAlignment.Center
            };
            Grid.SetColumn(unitLabel, 3);
            rowGrid.Children.Add(unitLabel);

            // Remove button
            var removeBtn = MakeButton("−",
                RedBtn, Colors.White, 30);
            removeBtn.Height = 30;
            removeBtn.FontSize = 16;
            removeBtn.FontWeight = FontWeights.Bold;
            removeBtn.Tag = rowGrid;
            removeBtn.Click += (s, e) =>
            {
                _rowsPanel.Children.Remove(rowGrid);
            };
            Grid.SetColumn(removeBtn, 4);
            rowGrid.Children.Add(removeBtn);

            // Combo selection changed → fill placeholder
            combo.SelectionChanged += (s, e) =>
            {
                string paramName = combo.SelectedItem as string;
                if (paramName == null) return;

                ParamInfo info = paramList
                    .FirstOrDefault(p => p.Name == paramName);
                if (info == null) return;

                unitLabel.Text = info.DisplayUnit;

                // Show current value as grey placeholder
                if (string.IsNullOrEmpty(valBox.Text))
                {
                    valBox.Foreground =
                        new SolidColorBrush(PlaceholderGray);
                    valBox.Text = info.CurrentValue;
                    valBox.Tag = "placeholder";
                }
            };

            // Clear placeholder on focus
            valBox.GotFocus += (s, e) =>
            {
                if (valBox.Tag as string == "placeholder")
                {
                    valBox.Text = "";
                    valBox.Foreground =
                        new SolidColorBrush(DarkText);
                    valBox.Tag = null;
                }
                valBorder.BorderBrush =
                    new SolidColorBrush(BluePrimary);
            };
            valBox.LostFocus += (s, e) =>
            {
                valBorder.BorderBrush =
                    new SolidColorBrush(BorderClr);
                if (string.IsNullOrEmpty(valBox.Text))
                {
                    string paramName =
                        combo.SelectedItem as string;
                    ParamInfo info = paramName != null
                        ? paramList.FirstOrDefault(
                            p => p.Name == paramName)
                        : null;
                    if (info != null)
                    {
                        valBox.Foreground =
                            new SolidColorBrush(PlaceholderGray);
                        valBox.Text = info.CurrentValue;
                        valBox.Tag = "placeholder";
                    }
                }
            };

            // If we already have a param selected, trigger
            if (combo.SelectedItem != null && valueText == "")
            {
                ParamInfo info = paramList
                    .FirstOrDefault(p =>
                        p.Name == (string)combo.SelectedItem);
                if (info != null)
                {
                    unitLabel.Text = info.DisplayUnit;
                    valBox.Foreground =
                        new SolidColorBrush(PlaceholderGray);
                    valBox.Text = info.CurrentValue;
                    valBox.Tag = "placeholder";
                }
            }

            _rowsPanel.Children.Add(rowGrid);
        }

        // ═══════════════════════════════════════════════════════
        // Apply
        // ═══════════════════════════════════════════════════════

        private void OnApply(object sender, RoutedEventArgs e)
        {
            SaveCurrentRows();

            ResultChanges =
                new Dictionary<string, List<ParamChange>>();

            foreach (string key in _familyKeys)
            {
                var rows = _familyRows[key];
                var changes = new List<ParamChange>();

                foreach (var rd in rows)
                {
                    if (string.IsNullOrEmpty(rd.SelectedParam))
                        continue;
                    if (string.IsNullOrEmpty(rd.ValueText))
                        continue;

                    // Skip if still showing placeholder
                    // (user didn't actually edit)
                    List<ParamInfo> pList = _familyParams[key];
                    ParamInfo info = pList.FirstOrDefault(
                        p => p.Name == rd.SelectedParam);
                    if (info != null
                        && rd.ValueText == info.CurrentValue)
                        continue;

                    changes.Add(new ParamChange
                    {
                        ParamName = rd.SelectedParam,
                        NewValue = rd.ValueText
                    });
                }

                if (changes.Count > 0)
                    ResultChanges[key] = changes;
            }

            if (ResultChanges.Count == 0)
            {
                MessageBox.Show(
                    "No changes detected. Edit the values "
                  + "before applying.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            DialogResult = true;
            Close();
        }

        // ═══════════════════════════════════════════════════════
        // Helpers
        // ═══════════════════════════════════════════════════════

        private T FindChild<T>(Grid grid) where T : class
        {
            foreach (UIElement child in grid.Children)
            {
                if (child is T t) return t;
                if (child is Border b && b.Child is T t2)
                    return t2;
            }
            return null;
        }

        private Button MakeButton(string text,
            Color bg, Color fg, double w)
        {
            var btn = new Button
            {
                Content = text,
                Width = w,
                Height = 34,
                FontSize = 12,
                Margin = new Thickness(0, 0, 8, 0),
                Cursor = Cursors.Hand,
                Foreground = new SolidColorBrush(fg),
                Background = new SolidColorBrush(bg),
                BorderThickness = new Thickness(0)
            };
            var tp = new ControlTemplate(typeof(Button));
            var bd = new FrameworkElementFactory(typeof(Border));
            bd.SetValue(Border.CornerRadiusProperty,
                new CornerRadius(6));
            bd.SetValue(Border.BackgroundProperty,
                new SolidColorBrush(bg));
            bd.SetValue(Border.PaddingProperty,
                new Thickness(10, 6, 10, 6));
            var cp = new FrameworkElementFactory(
                typeof(ContentPresenter));
            cp.SetValue(
                ContentPresenter.HorizontalAlignmentProperty,
                HorizontalAlignment.Center);
            cp.SetValue(
                ContentPresenter.VerticalAlignmentProperty,
                VerticalAlignment.Center);
            bd.AppendChild(cp);
            tp.VisualTree = bd;
            btn.Template = tp;
            return btn;
        }
    }
}
