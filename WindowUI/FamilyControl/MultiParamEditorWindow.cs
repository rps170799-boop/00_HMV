using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using Autodesk.Revit.DB;

using Color = System.Windows.Media.Color;
using Grid = System.Windows.Controls.Grid;
using HorizontalAlignment = System.Windows.HorizontalAlignment;
using VerticalAlignment = System.Windows.VerticalAlignment;

namespace HMVTools
{
    public class MultiParamEditorWindow : Window
    {
        // ── Data ──
        private Dictionary<string, List<ParamInfo>>
            _familyParams;
        private readonly Dictionary<string, List<FamilyInstance>>
            _familyGroups;
        private readonly bool _commonMode;
        private readonly List<string> _familyKeys;
        private readonly Document _doc;
        private readonly List<FamilyInstance> _allInstances;
        private readonly Func<Dictionary<string, List<ParamInfo>>>
            _refreshParams;

        private string _activeFamilyKey;

        // Per-family row states: familyKey → (paramName → state)
        private readonly Dictionary<string,
            Dictionary<string, ParamRowState>> _rowStates;

        // ── UI ──
        private StackPanel _rowsPanel;
        private StackPanel _radioPanel;
        private CheckBox _selectAllCheckBox;
        private TextBox _findBox;
        private TextBox _replaceBox;
        private TextBlock _statusLabel;
        private bool _suppressCheckSync;

        // ── Colors ──
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
        private static readonly Color OrangeAccent =
            Color.FromRgb(200, 120, 20);

        private class ParamRowState
        {
            public string NewValue = "";
            public bool IsChecked;
        }

        // ═════════════════════════════════════════════════════════
        // Constructor
        // ═════════════════════════════════════════════════════════

        public MultiParamEditorWindow(
            Dictionary<string, List<ParamInfo>> familyParams,
            Dictionary<string, List<FamilyInstance>> familyGroups,
            bool commonMode,
            Document doc,
            List<FamilyInstance> allInstances,
            Func<Dictionary<string, List<ParamInfo>>> refreshParams)
        {
            _familyParams = familyParams;
            _familyGroups = familyGroups;
            _commonMode = commonMode;
            _familyKeys = familyParams.Keys.ToList();
            _doc = doc;
            _allInstances = allInstances;
            _refreshParams = refreshParams;

            _rowStates = new Dictionary<string,
                Dictionary<string, ParamRowState>>();
            foreach (string key in _familyKeys)
                _rowStates[key] =
                    new Dictionary<string, ParamRowState>();

            _activeFamilyKey = _familyKeys[0];

            BuildUI();
        }

        // ═════════════════════════════════════════════════════════
        // Build UI
        // ═════════════════════════════════════════════════════════

        private void BuildUI()
        {
            Title = "HMV Tools – Multi Instance-Parameter Editor";
            Width = 750;
            MinWidth = 620;
            Height = 600;
            MinHeight = 460;
            WindowStartupLocation =
                WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.CanResize;
            Background = new SolidColorBrush(WindowBg);

            var root = new Grid
            {
                Margin = new Thickness(24)
            };
            root.RowDefinitions.Add(new RowDefinition
            { Height = GridLength.Auto });          // 0 title
            root.RowDefinitions.Add(new RowDefinition
            { Height = GridLength.Auto });          // 1 family selector
            root.RowDefinitions.Add(new RowDefinition
            { Height = GridLength.Auto });          // 2 column headers
            root.RowDefinitions.Add(new RowDefinition
            { Height = new GridLength(1, GridUnitType.Star) }); // 3 rows
            root.RowDefinitions.Add(new RowDefinition
            { Height = GridLength.Auto });          // 4 replace tool
            root.RowDefinitions.Add(new RowDefinition
            { Height = GridLength.Auto });          // 5 status bar
            root.RowDefinitions.Add(new RowDefinition
            { Height = GridLength.Auto });          // 6 buttons
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

            // ── Row 1: Family selector ──
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
                        Foreground =
                            new SolidColorBrush(DarkText),
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

            // ── Row 2: Column headers ──
            var headerGrid = MakeRowGrid();
            headerGrid.Margin = new Thickness(0, 0, 0, 4);

            _selectAllCheckBox = new CheckBox
            {
                IsChecked = false,
                HorizontalAlignment =
                    HorizontalAlignment.Center,
                VerticalAlignment =
                    VerticalAlignment.Center
            };
            _selectAllCheckBox.Checked += SelectAll_Changed;
            _selectAllCheckBox.Unchecked += SelectAll_Changed;
            Grid.SetColumn(_selectAllCheckBox, 0);
            headerGrid.Children.Add(_selectAllCheckBox);

            headerGrid.Children.Add(
                MakeHeaderLabel("Parameter", 1));
            headerGrid.Children.Add(
                MakeHeaderLabel("Current Value", 2));
            headerGrid.Children.Add(
                MakeHeaderLabel("New Value", 4));
            headerGrid.Children.Add(
                MakeHeaderLabel("Unit", 5));

            Grid.SetRow(headerGrid, 2);
            root.Children.Add(headerGrid);

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

            // ── Row 4: Replace tool ──
            var replaceSection = BuildReplaceRow();
            Grid.SetRow(replaceSection, 4);
            root.Children.Add(replaceSection);

            // ── Row 5: Status bar ──
            var statusBorder = new Border
            {
                CornerRadius = new CornerRadius(4),
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(BorderClr),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(10, 6, 10, 6),
                Margin = new Thickness(0, 0, 0, 12)
            };
            _statusLabel = new TextBlock
            {
                Text = "Select parameters and enter new values, "
                     + "then click Apply.",
                FontSize = 11,
                Foreground = new SolidColorBrush(MutedText)
            };
            statusBorder.Child = _statusLabel;
            Grid.SetRow(statusBorder, 5);
            root.Children.Add(statusBorder);

            // ── Row 6: Action buttons ──
            var btnRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var closeBtn = MakeButton("Close",
                GrayBg, Color.FromRgb(60, 60, 60), 90);
            closeBtn.Click += (s, e) => Close();
            btnRow.Children.Add(closeBtn);

            var applyBtn = MakeButton("Apply Changes",
                GreenBtn, Colors.White, 160);
            applyBtn.FontWeight = FontWeights.SemiBold;
            applyBtn.Click += OnApply;
            btnRow.Children.Add(applyBtn);

            Grid.SetRow(btnRow, 6);
            root.Children.Add(btnRow);

            // Load initial rows
            LoadRowsForFamily();
        }

        /// <summary>
        /// Creates a Grid with the standard column layout
        /// used for both the header and each parameter row.
        /// </summary>
        private Grid MakeRowGrid()
        {
            var grid = new Grid();
            // Col 0: Checkbox (28)
            grid.ColumnDefinitions.Add(
                new ColumnDefinition
                { Width = new GridLength(28) });
            // Col 1: Param name (1.4*)
            grid.ColumnDefinitions.Add(
                new ColumnDefinition
                {
                    Width = new GridLength(1.4,
                        GridUnitType.Star)
                });
            // Col 2: Current value (0.9*)
            grid.ColumnDefinitions.Add(
                new ColumnDefinition
                {
                    Width = new GridLength(0.9,
                        GridUnitType.Star)
                });
            // Col 3: Gap (8)
            grid.ColumnDefinitions.Add(
                new ColumnDefinition
                { Width = new GridLength(8) });
            // Col 4: New value (1*)
            grid.ColumnDefinitions.Add(
                new ColumnDefinition
                {
                    Width = new GridLength(1,
                        GridUnitType.Star)
                });
            // Col 5: Unit (45)
            grid.ColumnDefinitions.Add(
                new ColumnDefinition
                { Width = new GridLength(45) });
            return grid;
        }

        private TextBlock MakeHeaderLabel(string text, int col)
        {
            var tb = new TextBlock
            {
                Text = text,
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(MutedText),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0, 0, 0)
            };
            Grid.SetColumn(tb, col);
            return tb;
        }

        // ═════════════════════════════════════════════════════════
        // Replace tool
        // ═════════════════════════════════════════════════════════

        private StackPanel BuildReplaceRow()
        {
            var stack = new StackPanel
            {
                Orientation = Orientation.Vertical,
                Margin = new Thickness(0, 0, 0, 12)
            };

            var findRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 6)
            };
            findRow.Children.Add(new TextBlock
            {
                Text = "Find:",
                FontSize = 13,
                Width = 60,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 4, 0)
            });
            var findBorder = MakeInputBorder(200);
            _findBox = MakeInputTextBox();
            findBorder.Child = _findBox;
            findRow.Children.Add(findBorder);
            stack.Children.Add(findRow);

            var replaceRow = new StackPanel
            {
                Orientation = Orientation.Horizontal
            };
            replaceRow.Children.Add(new TextBlock
            {
                Text = "Replace:",
                FontSize = 13,
                Width = 60,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 4, 0)
            });
            var replaceBorder = MakeInputBorder(200);
            _replaceBox = MakeInputTextBox();
            replaceBorder.Child = _replaceBox;
            replaceRow.Children.Add(replaceBorder);

            var replaceSelBtn = MakeButton("Replace Selected",
                BluePrimary, Colors.White, 130);
            replaceSelBtn.Margin = new Thickness(8, 0, 0, 0);
            replaceSelBtn.Click += ReplaceSelected_Click;
            replaceRow.Children.Add(replaceSelBtn);

            var replaceAllBtn = MakeButton("Replace All",
                OrangeAccent, Colors.White, 100);
            replaceAllBtn.Margin = new Thickness(4, 0, 0, 0);
            replaceAllBtn.Click += ReplaceAll_Click;
            replaceRow.Children.Add(replaceAllBtn);

            stack.Children.Add(replaceRow);
            return stack;
        }

        // ═════════════════════════════════════════════════════════
        // Replace handlers
        // ═════════════════════════════════════════════════════════

        private void ReplaceSelected_Click(
            object sender, RoutedEventArgs e)
        {
            string find = _findBox.Text;
            if (string.IsNullOrEmpty(find))
            {
                MessageBox.Show("Enter the text to find.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }
            int count = DoReplace(find,
                _replaceBox.Text ?? "", true);
            if (count == 0)
                MessageBox.Show(
                    "No selected rows contained the find text "
                  + "in their New Value field.",
                    "HMV Tools", MessageBoxButton.OK);
        }

        private void ReplaceAll_Click(
            object sender, RoutedEventArgs e)
        {
            string find = _findBox.Text;
            if (string.IsNullOrEmpty(find))
            {
                MessageBox.Show("Enter the text to find.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }
            int count = DoReplace(find,
                _replaceBox.Text ?? "", false);
            if (count == 0)
                MessageBox.Show(
                    "No rows contained the find text "
                  + "in their New Value field.",
                    "HMV Tools", MessageBoxButton.OK);
        }

        private int DoReplace(string find, string replace,
            bool selectedOnly)
        {
            int count = 0;

            foreach (UIElement child in _rowsPanel.Children)
            {
                var rowGrid = child as Grid;
                if (rowGrid == null) continue;

                if (selectedOnly)
                {
                    var chk = FindChildInColumn<CheckBox>(
                        rowGrid, 0);
                    if (chk == null || chk.IsChecked != true)
                        continue;
                }

                var valBox = FindChildInColumn<TextBox>(
                    rowGrid, 4);
                if (valBox == null) continue;

                string current = valBox.Text ?? "";
                if (!current.Contains(find)) continue;

                valBox.Text = current.Replace(find, replace);
                valBox.Foreground =
                    new SolidColorBrush(DarkText);
                valBox.Tag = null;
                count++;
            }

            return count;
        }

        // ═════════════════════════════════════════════════════════
        // Checkbox: Select all / per-row
        // ═════════════════════════════════════════════════════════

        private void SelectAll_Changed(
            object sender, RoutedEventArgs e)
        {
            if (_suppressCheckSync) return;
            bool check = _selectAllCheckBox.IsChecked == true;

            foreach (UIElement child in _rowsPanel.Children)
            {
                var rowGrid = child as Grid;
                if (rowGrid == null) continue;
                var chk = FindChildInColumn<CheckBox>(
                    rowGrid, 0);
                if (chk != null) chk.IsChecked = check;

                // Enable/disable new-value textbox
                var valBox = FindChildInColumn<TextBox>(
                    rowGrid, 4);
                if (valBox != null)
                    valBox.IsEnabled = check;
            }
        }

        private void RowCheckBox_Click(
            object sender, RoutedEventArgs e)
        {
            if (_suppressCheckSync) return;

            // Enable/disable the value textbox in same row
            var chk = sender as CheckBox;
            if (chk != null)
            {
                var rowGrid = chk.Parent as Grid;
                if (rowGrid != null)
                {
                    var valBox = FindChildInColumn<TextBox>(
                        rowGrid, 4);
                    if (valBox != null)
                        valBox.IsEnabled =
                            chk.IsChecked == true;
                }
            }

            _suppressCheckSync = true;
            UpdateSelectAllState();
            _suppressCheckSync = false;
        }

        private void UpdateSelectAllState()
        {
            var states = new List<bool>();
            foreach (UIElement child in _rowsPanel.Children)
            {
                var rowGrid = child as Grid;
                if (rowGrid == null) continue;
                var chk = FindChildInColumn<CheckBox>(
                    rowGrid, 0);
                if (chk != null)
                    states.Add(chk.IsChecked == true);
            }

            if (states.Count == 0) return;

            if (states.All(s => s))
                _selectAllCheckBox.IsChecked = true;
            else if (states.All(s => !s))
                _selectAllCheckBox.IsChecked = false;
            else
                _selectAllCheckBox.IsChecked = null;
        }

        // ═════════════════════════════════════════════════════════
        // Family switching
        // ═════════════════════════════════════════════════════════

        private void OnFamilyRadioChecked(
            object sender, RoutedEventArgs e)
        {
            var rb = sender as RadioButton;
            if (rb == null) return;
            string newKey = rb.Tag as string;
            if (newKey == _activeFamilyKey) return;

            SaveCurrentRowStates();
            _activeFamilyKey = newKey;
            LoadRowsForFamily();
        }

        /// <summary>
        /// Reads the current UI state of all visible rows
        /// and saves it into _rowStates for the active family.
        /// </summary>
        private void SaveCurrentRowStates()
        {
            var states = _rowStates[_activeFamilyKey];

            foreach (UIElement child in _rowsPanel.Children)
            {
                var rowGrid = child as Grid;
                if (rowGrid == null) continue;

                string paramName = rowGrid.Tag as string;
                if (paramName == null) continue;

                var chk = FindChildInColumn<CheckBox>(
                    rowGrid, 0);
                var valBox = FindChildInColumn<TextBox>(
                    rowGrid, 4);

                bool isChecked = chk != null
                    && chk.IsChecked == true;
                string newVal = "";
                if (valBox != null
                    && valBox.Tag as string != "placeholder")
                    newVal = valBox.Text ?? "";

                if (states.ContainsKey(paramName))
                {
                    states[paramName].IsChecked = isChecked;
                    states[paramName].NewValue = newVal;
                }
                else
                {
                    states[paramName] = new ParamRowState
                    {
                        IsChecked = isChecked,
                        NewValue = newVal
                    };
                }
            }
        }

        /// <summary>
        /// Rebuilds all rows for the active family using
        /// current ParamInfo data and saved row states.
        /// </summary>
        private void LoadRowsForFamily()
        {
            _rowsPanel.Children.Clear();

            List<ParamInfo> paramList =
                _familyParams[_activeFamilyKey];
            var states = _rowStates[_activeFamilyKey];

            foreach (ParamInfo pi in paramList)
            {
                ParamRowState state = null;
                if (states.ContainsKey(pi.Name))
                    state = states[pi.Name];

                AddParamRow(pi, state);
            }

            _suppressCheckSync = true;
            UpdateSelectAllState();
            _suppressCheckSync = false;
        }

        // ═════════════════════════════════════════════════════════
        // Row building
        // ═════════════════════════════════════════════════════════

        /// <summary>
        /// Creates a single parameter row showing checkbox,
        /// name, current value, editable new-value field,
        /// and unit label.
        /// </summary>
        private void AddParamRow(
            ParamInfo pi, ParamRowState state)
        {
            var rowGrid = MakeRowGrid();
            rowGrid.Margin = new Thickness(0, 0, 0, 4);
            rowGrid.Tag = pi.Name;

            bool isChecked = state != null && state.IsChecked;
            string newVal = state != null ? state.NewValue : "";

            // ── Col 0: Checkbox ──
            var rowCheckBox = new CheckBox
            {
                IsChecked = isChecked,
                HorizontalAlignment =
                    HorizontalAlignment.Center,
                VerticalAlignment =
                    VerticalAlignment.Center
            };
            rowCheckBox.Click += RowCheckBox_Click;
            Grid.SetColumn(rowCheckBox, 0);
            rowGrid.Children.Add(rowCheckBox);

            // ── Col 1: Parameter name ──
            var nameLabel = new TextBlock
            {
                Text = pi.Name,
                FontSize = 12,
                Foreground = new SolidColorBrush(DarkText),
                VerticalAlignment =
                    VerticalAlignment.Center,
                Margin = new Thickness(4, 0, 4, 0),
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            Grid.SetColumn(nameLabel, 1);
            rowGrid.Children.Add(nameLabel);

            // ── Col 2: Current value ──
            var currentLabel = new TextBlock
            {
                Text = pi.CurrentValue,
                FontSize = 11,
                VerticalAlignment =
                    VerticalAlignment.Center,
                Margin = new Thickness(4, 0, 4, 0),
                TextTrimming = TextTrimming.CharacterEllipsis,
                ToolTip = pi.CurrentValue
            };
            if (pi.Varies)
            {
                currentLabel.Foreground =
                    new SolidColorBrush(OrangeAccent);
                currentLabel.FontStyle = FontStyles.Italic;
                currentLabel.ToolTip =
                    "Varies: " + pi.CurrentValue;
            }
            else
            {
                currentLabel.Foreground =
                    new SolidColorBrush(MutedText);
            }
            Grid.SetColumn(currentLabel, 2);
            rowGrid.Children.Add(currentLabel);

            // ── Col 4: New value textbox ──
            var valBorder = new Border
            {
                CornerRadius = new CornerRadius(6),
                BorderBrush = new SolidColorBrush(BorderClr),
                BorderThickness = new Thickness(1),
                Background = Brushes.White
            };
            var valBox = new TextBox
            {
                Height = 30,
                FontSize = 12,
                VerticalContentAlignment =
                    VerticalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(8, 0, 8, 0),
                IsEnabled = isChecked
            };

            // Set initial value: user-edited text or
            // placeholder with current value
            if (!string.IsNullOrEmpty(newVal))
            {
                valBox.Text = newVal;
                valBox.Foreground =
                    new SolidColorBrush(DarkText);
                valBox.Tag = null;
            }
            else
            {
                valBox.Foreground =
                    new SolidColorBrush(PlaceholderGray);
                valBox.Text = pi.Varies
                    ? "<varies>" : pi.CurrentValue;
                valBox.Tag = "placeholder";
            }

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

            // Restore placeholder on lost focus if empty
            valBox.LostFocus += (s, e) =>
            {
                valBorder.BorderBrush =
                    new SolidColorBrush(BorderClr);
                if (string.IsNullOrEmpty(valBox.Text))
                {
                    valBox.Foreground =
                        new SolidColorBrush(PlaceholderGray);
                    valBox.Text = pi.Varies
                        ? "<varies>" : pi.CurrentValue;
                    valBox.Tag = "placeholder";
                }
            };

            valBorder.Child = valBox;
            Grid.SetColumn(valBorder, 4);
            rowGrid.Children.Add(valBorder);

            // ── Col 5: Unit label ──
            var unitLabel = new TextBlock
            {
                Text = pi.DisplayUnit,
                FontSize = 10,
                Foreground = new SolidColorBrush(MutedText),
                VerticalAlignment =
                    VerticalAlignment.Center,
                HorizontalAlignment =
                    HorizontalAlignment.Center
            };
            Grid.SetColumn(unitLabel, 5);
            rowGrid.Children.Add(unitLabel);

            _rowsPanel.Children.Add(rowGrid);
        }

        // ═════════════════════════════════════════════════════════
        // Apply (live – no window close)
        // ═════════════════════════════════════════════════════════

        private void OnApply(object sender, RoutedEventArgs e)
        {
            SaveCurrentRowStates();

            // Collect changes across all families
            var allChanges = new Dictionary<string,
                List<ParamChange>>();

            foreach (string key in _familyKeys)
            {
                var states = _rowStates[key];
                var changes = new List<ParamChange>();

                foreach (var kvp in states)
                {
                    if (!kvp.Value.IsChecked) continue;
                    if (string.IsNullOrWhiteSpace(
                        kvp.Value.NewValue)) continue;

                    changes.Add(new ParamChange
                    {
                        ParamName = kvp.Key,
                        NewValue = kvp.Value.NewValue
                    });
                }

                if (changes.Count > 0)
                    allChanges[key] = changes;
            }

            if (allChanges.Count == 0)
            {
                _statusLabel.Text =
                    "No changes to apply. Check parameters "
                  + "and enter new values first.";
                _statusLabel.Foreground =
                    new SolidColorBrush(OrangeAccent);
                return;
            }

            int changed = 0;
            int errors = 0;

            using (Transaction tx = new Transaction(_doc,
                "HMV – Multi-Parameter Edit"))
            {
                tx.Start();

                foreach (var entry in allChanges)
                {
                    string famKey = entry.Key;
                    List<FamilyInstance> targets;

                    if (_commonMode)
                    {
                        targets = _allInstances;
                    }
                    else
                    {
                        if (!_familyGroups.ContainsKey(famKey))
                            continue;
                        targets = _familyGroups[famKey];
                    }

                    foreach (var change in entry.Value)
                    {
                        foreach (FamilyInstance fi in targets)
                        {
                            Parameter p = fi.LookupParameter(
                                change.ParamName);
                            if (p == null || p.IsReadOnly)
                                continue;

                            try
                            {
                                if (p.StorageType ==
                                    StorageType.Double)
                                {
                                    double val = double.Parse(
                                        change.NewValue);
                                    try
                                    {
                                        ForgeTypeId specId =
                                            p.Definition
                                             .GetDataType();
                                        ForgeTypeId unitId =
                                            _doc.GetUnits()
                                            .GetFormatOptions(
                                                specId)
                                            .GetUnitTypeId();
                                        val = UnitUtils
                                            .ConvertToInternalUnits(
                                                val, unitId);
                                    }
                                    catch { }
                                    p.Set(val);
                                    changed++;
                                }
                                else if (p.StorageType ==
                                    StorageType.Integer)
                                {
                                    if (int.TryParse(
                                        change.NewValue,
                                        out int iv))
                                    {
                                        p.Set(iv);
                                        changed++;
                                    }
                                    else errors++;
                                }
                                else if (p.StorageType ==
                                    StorageType.String)
                                {
                                    p.Set(change.NewValue);
                                    changed++;
                                }
                            }
                            catch
                            {
                                errors++;
                            }
                        }
                    }
                }

                tx.Commit();
            }

            // Refresh ParamInfo data from the model
            _familyParams = _refreshParams();

            // Clear new-value text for applied params
            // but keep checkbox state
            foreach (string key in _familyKeys)
            {
                var states = _rowStates[key];
                foreach (var kvp in states)
                    kvp.Value.NewValue = "";
            }

            // Rebuild UI so current values are up to date
            LoadRowsForFamily();

            // Update status
            string msg = "Applied: " + changed
                + " value(s) set";
            if (errors > 0)
                msg += ", " + errors + " error(s)";

            _statusLabel.Text = msg;
            _statusLabel.Foreground = errors > 0
                ? new SolidColorBrush(RedBtn)
                : new SolidColorBrush(GreenBtn);
        }

        // ═════════════════════════════════════════════════════════
        // Helpers
        // ═════════════════════════════════════════════════════════

        private T FindChildInColumn<T>(Grid grid, int column)
            where T : class
        {
            foreach (UIElement child in grid.Children)
            {
                if (Grid.GetColumn(child) != column) continue;
                if (child is T t) return t;
                if (child is Border b && b.Child is T t2)
                    return t2;
            }
            return null;
        }

        private Border MakeInputBorder(double width)
        {
            return new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderClr),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Width = width
            };
        }

        private TextBox MakeInputTextBox()
        {
            return new TextBox
            {
                Height = 32,
                FontSize = 13,
                VerticalContentAlignment =
                    VerticalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(10, 0, 10, 0)
            };
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
            var bd = new FrameworkElementFactory(
                typeof(Border));
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