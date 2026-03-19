using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

using Color = System.Windows.Media.Color;
using HorizontalAlignment = System.Windows.HorizontalAlignment;
using VerticalAlignment = System.Windows.VerticalAlignment;

namespace HMVTools
{
    // ── Data class (plain, no Revit references) ────────────────

    public class ViewAuditEntry : INotifyPropertyChanged
    {
        public int ElementId { get; set; }
        public string ViewType { get; set; }
        public string OriginalName { get; set; }

        private string _newName;
        public string NewName
        {
            get => _newName;
            set
            {
                if (_newName != value)
                {
                    _newName = value;
                    OnPropertyChanged(nameof(NewName));
                    OnPropertyChanged(nameof(NameChanged));
                }
            }
        }

        public string Sheets { get; set; }
        public int SheetCount { get; set; }

        private bool _isSelected = true;
        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); }
        }

        private bool _hasConflict;
        public bool HasConflict
        {
            get => _hasConflict;
            set { _hasConflict = value; OnPropertyChanged(nameof(HasConflict)); }
        }

        public bool NameChanged => NewName != OriginalName;

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // ── Conflict highlight converter ───────────────────────────

    public class ConflictBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType,
            object parameter, CultureInfo culture)
        {
            if (value is bool conflict && conflict)
                return new SolidColorBrush(Color.FromArgb(55, 220, 50, 50));
            return Brushes.Transparent;
        }

        public object ConvertBack(object value, Type targetType,
            object parameter, CultureInfo culture) =>
            throw new NotImplementedException();
    }

    // ── Changed-name highlight converter ───────────────────────

    public class ChangedForegroundConverter : IMultiValueConverter
    {
        private static readonly SolidColorBrush Normal =
            new SolidColorBrush(Color.FromRgb(30, 30, 30));
        private static readonly SolidColorBrush Changed =
            new SolidColorBrush(Color.FromRgb(0, 100, 180));
        private static readonly SolidColorBrush Conflict =
            new SolidColorBrush(Color.FromRgb(180, 30, 30));

        public object Convert(object[] values, Type targetType,
            object parameter, CultureInfo culture)
        {
            bool hasConflict = values.Length > 1
                && values[1] is bool c && c;
            bool nameChanged = values.Length > 0
                && values[0] is string newName
                && values.Length > 2
                && values[2] is string origName
                && newName != origName;

            if (hasConflict) return Conflict;
            if (nameChanged) return Changed;
            return Normal;
        }

        public object[] ConvertBack(object value, Type[] targetTypes,
            object parameter, CultureInfo culture) =>
            throw new NotImplementedException();
    }

    // ── Window ─────────────────────────────────────────────────

    public class ViewAuditWindow : Window
    {
        // Colors (matching HMV Tools palette)
        private static readonly Color BluePrimary = Color.FromRgb(0, 120, 212);
        private static readonly Color GrayBg = Color.FromRgb(240, 240, 243);
        private static readonly Color DarkText = Color.FromRgb(30, 30, 30);
        private static readonly Color MutedText = Color.FromRgb(120, 120, 130);
        private static readonly Color BorderColor = Color.FromRgb(200, 200, 210);
        private static readonly Color WindowBg = Color.FromRgb(245, 245, 248);
        private static readonly Color GreenAccent = Color.FromRgb(40, 160, 80);

        // Controls
        private DataGrid dataGrid;
        private TextBox searchBox;
        private TextBox prefixBox;
        private TextBlock summaryText;
        private CheckBox selectAllCheckBox;

        // Data
        private ObservableCollection<ViewAuditEntry> allEntries;
        private ICollectionView collectionView;
        private HashSet<string> externalViewNames;

        /// <summary>Filled when user clicks Apply. Null if cancelled.</summary>
        public List<ViewAuditEntry> Results { get; private set; }

        /// <param name="entries">Views placed on sheets.</param>
        /// <param name="allProjectViewNames">
        /// Every view name in the project, for conflict detection.</param>
        public ViewAuditWindow(
            List<ViewAuditEntry> entries,
            HashSet<string> allProjectViewNames)
        {
            // Build set of names NOT in our table (external)
            var tableNames = new HashSet<string>(
                entries.Select(e => e.OriginalName));
            externalViewNames = new HashSet<string>(
                allProjectViewNames.Where(n => !tableNames.Contains(n)));

            allEntries = new ObservableCollection<ViewAuditEntry>(entries);

            // Subscribe to each entry for live conflict detection
            foreach (var entry in allEntries)
                entry.PropertyChanged += Entry_PropertyChanged;

            Title = "HMV Tools – View Audit";
            Width = 960;
            Height = 680;
            MinWidth = 780;
            MinHeight = 520;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.CanResize;
            Background = new SolidColorBrush(WindowBg);

            var mainGrid = new Grid { Margin = new Thickness(20) };
            mainGrid.RowDefinitions.Add(Row(GridLength.Auto));   // 0 Title
            mainGrid.RowDefinitions.Add(Row(GridLength.Auto));   // 1 Search + summary
            mainGrid.RowDefinitions.Add(Row(Star()));            // 2 DataGrid
            mainGrid.RowDefinitions.Add(Row(GridLength.Auto));   // 3 Prefix row
            mainGrid.RowDefinitions.Add(Row(GridLength.Auto));   // 4 Buttons

            // ── Row 0: Title ───────────────────────────────────
            var title = new TextBlock
            {
                Text = "View Audit – Sheets",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 4)
            };
            Grid.SetRow(title, 0);
            mainGrid.Children.Add(title);

            // ── Row 1: Search bar + summary ────────────────────
            var topRow = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            topRow.ColumnDefinitions.Add(
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            topRow.ColumnDefinitions.Add(
                new ColumnDefinition { Width = GridLength.Auto });

            // Search
            var searchBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 12, 0)
            };
            searchBox = new TextBox
            {
                Height = 30,
                FontSize = 13,
                VerticalContentAlignment = VerticalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(10, 0, 10, 0)
            };
            searchBox.TextChanged += SearchBox_TextChanged;
            searchBox.GotFocus += (s, e) =>
                searchBorder.BorderBrush = new SolidColorBrush(BluePrimary);
            searchBox.LostFocus += (s, e) =>
                searchBorder.BorderBrush = new SolidColorBrush(BorderColor);
            searchBorder.Child = searchBox;
            Grid.SetColumn(searchBorder, 0);
            topRow.Children.Add(searchBorder);

            // Summary
            summaryText = new TextBlock
            {
                FontSize = 12,
                Foreground = new SolidColorBrush(MutedText),
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(summaryText, 1);
            topRow.Children.Add(summaryText);

            Grid.SetRow(topRow, 1);
            mainGrid.Children.Add(topRow);

            // ── Row 2: DataGrid ────────────────────────────────
            dataGrid = BuildDataGrid();
            collectionView = CollectionViewSource.GetDefaultView(allEntries);
            dataGrid.ItemsSource = collectionView;

            var gridBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 10),
                Child = dataGrid
            };
            Grid.SetRow(gridBorder, 2);
            mainGrid.Children.Add(gridBorder);

            // ── Row 3: Prefix row ──────────────────────────────
            var prefixRow = BuildPrefixRow();
            Grid.SetRow(prefixRow, 3);
            mainGrid.Children.Add(prefixRow);

            // ── Row 4: Action buttons ──────────────────────────
            var buttonPanel = BuildButtonRow();
            Grid.SetRow(buttonPanel, 4);
            mainGrid.Children.Add(buttonPanel);

            Content = mainGrid;

            // Initial state
            UpdateSummary();
            CheckConflicts();

            Loaded += (s, e) => searchBox.Focus();
        }

        // ── DataGrid setup ─────────────────────────────────────

        private DataGrid BuildDataGrid()
        {
            var dg = new DataGrid
            {
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                CanUserReorderColumns = false,
                CanUserSortColumns = true,
                SelectionUnit = DataGridSelectionUnit.FullRow,
                SelectionMode = DataGridSelectionMode.Extended,
                GridLinesVisibility = DataGridGridLinesVisibility.Horizontal,
                HorizontalGridLinesBrush =
                    new SolidColorBrush(Color.FromRgb(235, 235, 240)),
                BorderThickness = new Thickness(0),
                Background = Brushes.White,
                RowHeaderWidth = 0,
                FontSize = 12.5,
                IsReadOnly = false
            };

            // Alternating row colors
            dg.AlternatingRowBackground =
                new SolidColorBrush(Color.FromRgb(250, 250, 253));

            // Header style
            var headerStyle = new Style(typeof(DataGridColumnHeader));
            headerStyle.Setters.Add(new Setter(
                DataGridColumnHeader.BackgroundProperty,
                new SolidColorBrush(GrayBg)));
            headerStyle.Setters.Add(new Setter(
                DataGridColumnHeader.PaddingProperty,
                new Thickness(8, 7, 8, 7)));
            headerStyle.Setters.Add(new Setter(
                DataGridColumnHeader.FontWeightProperty,
                FontWeights.SemiBold));
            headerStyle.Setters.Add(new Setter(
                DataGridColumnHeader.FontSizeProperty, 12.0));
            headerStyle.Setters.Add(new Setter(
                DataGridColumnHeader.ForegroundProperty,
                new SolidColorBrush(Color.FromRgb(60, 60, 70))));
            headerStyle.Setters.Add(new Setter(
                DataGridColumnHeader.BorderBrushProperty,
                new SolidColorBrush(Color.FromRgb(220, 220, 228))));
            headerStyle.Setters.Add(new Setter(
                DataGridColumnHeader.BorderThicknessProperty,
                new Thickness(0, 0, 1, 1)));
            dg.ColumnHeaderStyle = headerStyle;

            // ── Column 0: Checkbox ─────────────────────────────
            var chkCol = new DataGridTemplateColumn
            {
                Width = new DataGridLength(36)
            };

            // Header: Select-all checkbox
            var headerFactory = new FrameworkElementFactory(typeof(CheckBox));
            selectAllCheckBox = new CheckBox
            {
                IsChecked = true,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            selectAllCheckBox.Checked += SelectAll_Changed;
            selectAllCheckBox.Unchecked += SelectAll_Changed;
            chkCol.Header = selectAllCheckBox;

            // Cell: per-row checkbox
            var chkFactory = new FrameworkElementFactory(typeof(CheckBox));
            chkFactory.SetBinding(
                System.Windows.Controls.Primitives.ToggleButton.IsCheckedProperty,
                new Binding("IsSelected")
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                });
            chkFactory.SetValue(HorizontalAlignmentProperty,
                HorizontalAlignment.Center);
            chkFactory.SetValue(VerticalAlignmentProperty,
                VerticalAlignment.Center);
            var chkTemplate = new DataTemplate { VisualTree = chkFactory };
            chkCol.CellTemplate = chkTemplate;
            dg.Columns.Add(chkCol);

            // ── Column 1: View Type (read-only) ────────────────
            var typeCol = new DataGridTextColumn
            {
                Header = "Type",
                Binding = new Binding("ViewType"),
                Width = new DataGridLength(110),
                IsReadOnly = true
            };
            dg.Columns.Add(typeCol);

            // ── Column 2: Original Name (read-only) ────────────
            var origCol = new DataGridTextColumn
            {
                Header = "Original Name",
                Binding = new Binding("OriginalName"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                IsReadOnly = true
            };
            var origCellStyle = new Style(typeof(DataGridCell));
            origCellStyle.Setters.Add(new Setter(
                DataGridCell.ForegroundProperty,
                new SolidColorBrush(MutedText)));
            origCol.CellStyle = origCellStyle;
            dg.Columns.Add(origCol);

            // ── Column 3: New Name (editable) ──────────────────
            var newNameCol = new DataGridTemplateColumn
            {
                Header = "New Name",
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                SortMemberPath = "NewName"
            };

            // Display template
            var displayFactory =
                new FrameworkElementFactory(typeof(TextBlock));
            displayFactory.SetBinding(TextBlock.TextProperty,
                new Binding("NewName"));
            displayFactory.SetValue(TextBlock.PaddingProperty,
                new Thickness(6, 4, 6, 4));
            displayFactory.SetValue(TextBlock.VerticalAlignmentProperty,
                VerticalAlignment.Center);
            var displayTemplate = new DataTemplate
            { VisualTree = displayFactory };
            newNameCol.CellTemplate = displayTemplate;

            // Editing template
            var editFactory =
                new FrameworkElementFactory(typeof(TextBox));
            editFactory.SetBinding(TextBox.TextProperty,
                new Binding("NewName")
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                });
            editFactory.SetValue(TextBox.BorderThicknessProperty,
                new Thickness(0));
            editFactory.SetValue(TextBox.PaddingProperty,
                new Thickness(4, 2, 4, 2));
            editFactory.SetValue(TextBox.FontSizeProperty, 12.5);
            var editTemplate = new DataTemplate
            { VisualTree = editFactory };
            newNameCol.CellEditingTemplate = editTemplate;

            // Conflict cell style – red background
            var nameCellStyle = new Style(typeof(DataGridCell));
            var conflictTrigger = new DataTrigger
            {
                Binding = new Binding("HasConflict"),
                Value = true
            };
            conflictTrigger.Setters.Add(new Setter(
                DataGridCell.BackgroundProperty,
                new SolidColorBrush(Color.FromArgb(55, 220, 50, 50))));
            conflictTrigger.Setters.Add(new Setter(
                DataGridCell.ToolTipProperty,
                "Duplicate name – this view will be skipped"));
            nameCellStyle.Triggers.Add(conflictTrigger);

            // Changed name: blue foreground
            var changedTrigger = new DataTrigger
            {
                Binding = new Binding("NameChanged"),
                Value = true
            };
            changedTrigger.Setters.Add(new Setter(
                DataGridCell.ForegroundProperty,
                new SolidColorBrush(Color.FromRgb(0, 100, 180))));
            nameCellStyle.Triggers.Add(changedTrigger);

            newNameCol.CellStyle = nameCellStyle;
            dg.Columns.Add(newNameCol);

            // ── Column 4: Sheets (read-only) ───────────────────
            var sheetCol = new DataGridTextColumn
            {
                Header = "Sheet(s)",
                Binding = new Binding("Sheets"),
                Width = new DataGridLength(200),
                IsReadOnly = true
            };
            var sheetCellStyle = new Style(typeof(DataGridCell));
            sheetCellStyle.Setters.Add(new Setter(
                DataGridCell.ForegroundProperty,
                new SolidColorBrush(MutedText)));
            sheetCellStyle.Setters.Add(new Setter(
                DataGridCell.FontSizeProperty, 11.5));
            sheetCol.CellStyle = sheetCellStyle;
            dg.Columns.Add(sheetCol);

            // Row style for conflict highlight on entire row
            var rowStyle = new Style(typeof(DataGridRow));
            var rowConflict = new DataTrigger
            {
                Binding = new Binding("HasConflict"),
                Value = true
            };
            rowConflict.Setters.Add(new Setter(
                DataGridRow.BackgroundProperty,
                new SolidColorBrush(Color.FromArgb(25, 220, 50, 50))));
            rowStyle.Triggers.Add(rowConflict);
            dg.RowStyle = rowStyle;

            // Handle edit commit for conflict check
            dg.CellEditEnding += DataGrid_CellEditEnding;

            return dg;
        }

        // ── Prefix row ─────────────────────────────────────────

        private Grid BuildPrefixRow()
        {
            var row = new Grid { Margin = new Thickness(0, 0, 0, 12) };
            row.ColumnDefinitions.Add(new ColumnDefinition
            { Width = GridLength.Auto });
            row.ColumnDefinitions.Add(new ColumnDefinition
            { Width = new GridLength(240) });
            row.ColumnDefinitions.Add(new ColumnDefinition
            { Width = GridLength.Auto });
            row.ColumnDefinitions.Add(new ColumnDefinition
            { Width = GridLength.Auto });
            row.ColumnDefinitions.Add(new ColumnDefinition
            { Width = new GridLength(1, GridUnitType.Star) });

            // Label
            var lbl = new TextBlock
            {
                Text = "Prefix:",
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 0)
            };
            Grid.SetColumn(lbl, 0);
            row.Children.Add(lbl);

            // TextBox
            var prefixBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 8, 0)
            };
            prefixBox = new TextBox
            {
                Height = 32,
                FontSize = 13,
                VerticalContentAlignment = VerticalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(10, 0, 10, 0)
            };
            prefixBox.GotFocus += (s, e) =>
                prefixBorder.BorderBrush = new SolidColorBrush(BluePrimary);
            prefixBox.LostFocus += (s, e) =>
                prefixBorder.BorderBrush = new SolidColorBrush(BorderColor);
            prefixBorder.Child = prefixBox;
            Grid.SetColumn(prefixBorder, 1);
            row.Children.Add(prefixBorder);

            // Add Prefix button
            var addBtn = CreateButton("Add Prefix", BluePrimary,
                Color.FromRgb(255, 255, 255));
            addBtn.Width = 110;
            addBtn.Click += AddPrefix_Click;
            Grid.SetColumn(addBtn, 2);
            row.Children.Add(addBtn);

            // Reset button
            var resetBtn = CreateButton("Reset Names", GrayBg,
                Color.FromRgb(80, 80, 80));
            resetBtn.Width = 110;
            resetBtn.Margin = new Thickness(8, 0, 0, 0);
            resetBtn.Click += ResetNames_Click;
            Grid.SetColumn(resetBtn, 3);
            row.Children.Add(resetBtn);

            return row;
        }

        // ── Action buttons row ─────────────────────────────────

        private StackPanel BuildButtonRow()
        {
            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var exitBtn = CreateButton("Exit", GrayBg,
                Color.FromRgb(60, 60, 60));
            exitBtn.Width = 90;
            exitBtn.Margin = new Thickness(0, 0, 8, 0);
            exitBtn.Click += (s, e) =>
            {
                DialogResult = false;
                Close();
            };

            var applyBtn = CreateButton("Apply Changes", GreenAccent,
                Color.FromRgb(255, 255, 255));
            applyBtn.Width = 150;
            applyBtn.Click += Apply_Click;

            panel.Children.Add(exitBtn);
            panel.Children.Add(applyBtn);
            return panel;
        }

        // ── Event handlers ─────────────────────────────────────

        private void Entry_PropertyChanged(object sender,
            PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "NewName")
            {
                CheckConflicts();
                UpdateSummary();
            }
        }

        private void DataGrid_CellEditEnding(object sender,
            DataGridCellEditEndingEventArgs e)
        {
            // Conflict check will fire via PropertyChanged,
            // but force a refresh after commit
            Dispatcher.BeginInvoke(new Action(() =>
            {
                CheckConflicts();
                UpdateSummary();
            }));
        }

        private void SelectAll_Changed(object sender, RoutedEventArgs e)
        {
            bool check = selectAllCheckBox.IsChecked == true;
            foreach (var entry in allEntries)
                entry.IsSelected = check;
        }

        private void SearchBox_TextChanged(object sender,
            TextChangedEventArgs e)
        {
            string filter = searchBox.Text.Trim().ToLower();
            collectionView.Filter = obj =>
            {
                if (string.IsNullOrEmpty(filter)) return true;
                var entry = obj as ViewAuditEntry;
                if (entry == null) return false;
                return entry.OriginalName.ToLower().Contains(filter)
                    || entry.NewName.ToLower().Contains(filter)
                    || entry.ViewType.ToLower().Contains(filter)
                    || entry.Sheets.ToLower().Contains(filter);
            };
        }

        private void AddPrefix_Click(object sender, RoutedEventArgs e)
        {
            string prefix = prefixBox.Text;
            if (string.IsNullOrEmpty(prefix))
            {
                MessageBox.Show("Enter a prefix first.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            int count = 0;
            foreach (var entry in allEntries)
            {
                if (!entry.IsSelected) continue;
                if (!entry.NewName.StartsWith(prefix))
                {
                    entry.NewName = prefix + entry.NewName;
                    count++;
                }
            }

            CheckConflicts();
            UpdateSummary();

            if (count == 0)
                MessageBox.Show(
                    "No views were modified. Either none are selected "
                    + "or all already have this prefix.",
                    "HMV Tools", MessageBoxButton.OK);
        }

        private void ResetNames_Click(object sender, RoutedEventArgs e)
        {
            foreach (var entry in allEntries)
                entry.NewName = entry.OriginalName;
            CheckConflicts();
            UpdateSummary();
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            CheckConflicts();

            int changes = allEntries.Count(x => x.NameChanged && !x.HasConflict);
            int conflicts = allEntries.Count(x => x.HasConflict);

            if (changes == 0 && conflicts == 0)
            {
                MessageBox.Show("No name changes to apply.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            string msg = $"{changes} view(s) will be renamed.";
            if (conflicts > 0)
                msg += $"\n{conflicts} view(s) have duplicate names "
                     + "and will be SKIPPED (shown in red).";
            msg += "\n\nProceed?";

            var result = MessageBox.Show(msg, "HMV Tools – Confirm",
                MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes) return;

            Results = allEntries.ToList();
            DialogResult = true;
            Close();
        }

        // ── Conflict detection ─────────────────────────────────

        private void CheckConflicts()
        {
            // Build frequency map of NewName among all entries
            var nameCounts = new Dictionary<string, int>(
                StringComparer.Ordinal);
            foreach (var entry in allEntries)
            {
                if (!nameCounts.ContainsKey(entry.NewName))
                    nameCounts[entry.NewName] = 0;
                nameCounts[entry.NewName]++;
            }

            foreach (var entry in allEntries)
            {
                bool conflict = false;

                // Duplicate within the table
                if (nameCounts.ContainsKey(entry.NewName)
                    && nameCounts[entry.NewName] > 1)
                    conflict = true;

                // Conflict with project views not in our table
                if (!conflict
                    && entry.NewName != entry.OriginalName
                    && externalViewNames.Contains(entry.NewName))
                    conflict = true;

                entry.HasConflict = conflict;
            }
        }

        // ── Summary update ─────────────────────────────────────

        private void UpdateSummary()
        {
            int total = allEntries.Count;
            int modified = allEntries.Count(e => e.NameChanged);
            int conflicts = allEntries.Count(e => e.HasConflict);

            summaryText.Text =
                $"Total: {total}    Modified: {modified}    "
                + $"Conflicts: {conflicts}";
        }

        // ── UI helpers (same pattern as PipeAnnotationWindow) ──

        private Button CreateButton(string text, Color bgColor,
            Color fgColor)
        {
            var btn = new Button
            {
                Content = text,
                Height = 36,
                FontSize = 13,
                Foreground = new SolidColorBrush(fgColor),
                Background = new SolidColorBrush(bgColor),
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
            btn.Template = GetRoundButtonTemplate(bgColor);
            return btn;
        }

        private ControlTemplate GetRoundButtonTemplate(Color bgColor)
        {
            var template = new ControlTemplate(typeof(Button));
            var border = new FrameworkElementFactory(typeof(Border));
            border.SetValue(Border.CornerRadiusProperty,
                new CornerRadius(6));
            border.SetValue(Border.BackgroundProperty,
                new SolidColorBrush(bgColor));
            border.SetValue(Border.PaddingProperty,
                new Thickness(14, 6, 14, 6));

            var content =
                new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(
                ContentPresenter.HorizontalAlignmentProperty,
                HorizontalAlignment.Center);
            content.SetValue(
                ContentPresenter.VerticalAlignmentProperty,
                VerticalAlignment.Center);

            border.AppendChild(content);
            template.VisualTree = border;
            return template;
        }

        private static RowDefinition Row(GridLength h) =>
            new RowDefinition { Height = h };

        private static GridLength Star() =>
            new GridLength(1, GridUnitType.Star);
    }
}