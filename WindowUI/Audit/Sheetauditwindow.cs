using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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

    public class SheetAuditEntry : INotifyPropertyChanged
    {
        public int ElementId { get; set; }
        public string OriginalNumber { get; set; }
        public string OriginalName { get; set; }

        private string _newNumber;
        public string NewNumber
        {
            get => _newNumber;
            set
            {
                if (_newNumber != value)
                {
                    _newNumber = value;
                    OnPropertyChanged(nameof(NewNumber));
                    OnPropertyChanged(nameof(NumberChanged));
                }
            }
        }

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

        public int ViewCount { get; set; }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); }
        }

        private bool _hasNumberConflict;
        public bool HasNumberConflict
        {
            get => _hasNumberConflict;
            set { _hasNumberConflict = value; OnPropertyChanged(nameof(HasNumberConflict)); }
        }

        public bool NumberChanged => NewNumber != OriginalNumber;
        public bool NameChanged => NewName != OriginalName;
        public bool AnyChanged => NumberChanged || NameChanged;

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // ── Window ─────────────────────────────────────────────────

    public class SheetAuditWindow : Window
    {
        // Colors
        private static readonly Color BluePrimary = Color.FromRgb(0, 120, 212);
        private static readonly Color GrayBg = Color.FromRgb(240, 240, 243);
        private static readonly Color DarkText = Color.FromRgb(30, 30, 30);
        private static readonly Color MutedText = Color.FromRgb(120, 120, 130);
        private static readonly Color BorderColor = Color.FromRgb(200, 200, 210);
        private static readonly Color WindowBg = Color.FromRgb(245, 245, 248);
        private static readonly Color GreenAccent = Color.FromRgb(40, 160, 80);
        private static readonly Color OrangeAccent = Color.FromRgb(200, 120, 20);

        // Controls
        private DataGrid dataGrid;
        private TextBox searchBox;
        private TextBox prefixNumBox;
        private TextBox prefixNameBox;
        private TextBox cutNumBox;
        private TextBox cutNameBox;
        private TextBlock summaryText;
        private CheckBox selectAllCheckBox;

        // Data
        private ObservableCollection<SheetAuditEntry> allEntries;
        private ICollectionView collectionView;
        private HashSet<string> externalSheetNumbers;

        private bool _suppressCheckSync;

        /// <summary>Filled when user clicks Apply. Null if cancelled.</summary>
        public List<SheetAuditEntry> Results { get; private set; }

        /// <param name="entries">All sheets in the project.</param>
        /// <param name="allProjectSheetNumbers">
        /// Every sheet number in the project (for duplicate detection).
        /// </param>
        public SheetAuditWindow(
            List<SheetAuditEntry> entries,
            HashSet<string> allProjectSheetNumbers)
        {
            // External = numbers NOT in our table
            var tableNumbers = new HashSet<string>(
                entries.Select(e => e.OriginalNumber));
            externalSheetNumbers = new HashSet<string>(
                allProjectSheetNumbers.Where(n => !tableNumbers.Contains(n)));

            allEntries = new ObservableCollection<SheetAuditEntry>(entries);

            foreach (var entry in allEntries)
                entry.PropertyChanged += Entry_PropertyChanged;

            Title = "HMV Tools \u2013 Sheet Audit";
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
            mainGrid.RowDefinitions.Add(Row(GridLength.Auto));   // 3 Tools
            mainGrid.RowDefinitions.Add(Row(GridLength.Auto));   // 4 Buttons

            // ── Row 0: Title ───────────────────────────────────
            var title = new TextBlock
            {
                Text = "Sheet Audit",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 4)
            };
            Grid.SetRow(title, 0);
            mainGrid.Children.Add(title);

            // ── Row 1: Search + summary ────────────────────────
            var topRow = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            topRow.ColumnDefinitions.Add(
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            topRow.ColumnDefinitions.Add(
                new ColumnDefinition { Width = GridLength.Auto });

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

            // ── Row 3: Tools ───────────────────────────────────
            var toolRow = BuildToolRow();
            Grid.SetRow(toolRow, 3);
            mainGrid.Children.Add(toolRow);

            // ── Row 4: Buttons ─────────────────────────────────
            var buttonPanel = BuildButtonRow();
            Grid.SetRow(buttonPanel, 4);
            mainGrid.Children.Add(buttonPanel);

            Content = mainGrid;

            UpdateSummary();
            CheckConflicts();

            Loaded += (s, e) => searchBox.Focus();
        }

        // ── DataGrid ───────────────────────────────────────────

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

            // ── Checkbox column ────────────────────────────────
            var chkCol = new DataGridTemplateColumn
            { Width = new DataGridLength(36) };

            selectAllCheckBox = new CheckBox
            {
                IsChecked = false,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            selectAllCheckBox.Checked += SelectAll_Changed;
            selectAllCheckBox.Unchecked += SelectAll_Changed;
            chkCol.Header = selectAllCheckBox;

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
            chkFactory.AddHandler(CheckBox.ClickEvent,
                new RoutedEventHandler(RowCheckBox_Click));
            chkCol.CellTemplate = new DataTemplate { VisualTree = chkFactory };
            dg.Columns.Add(chkCol);

            // ── Original Number (read-only) ────────────────────
            var origNumCol = new DataGridTextColumn
            {
                Header = "Original #",
                Binding = new Binding("OriginalNumber"),
                Width = new DataGridLength(100),
                IsReadOnly = true
            };
            var mutedStyle = new Style(typeof(DataGridCell));
            mutedStyle.Setters.Add(new Setter(
                DataGridCell.ForegroundProperty,
                new SolidColorBrush(MutedText)));
            origNumCol.CellStyle = mutedStyle;
            dg.Columns.Add(origNumCol);

            // ── New Number (editable) ──────────────────────────
            var newNumCol = new DataGridTemplateColumn
            {
                Header = "New #",
                Width = new DataGridLength(100),
                SortMemberPath = "NewNumber"
            };

            var numDisplayFactory = new FrameworkElementFactory(typeof(TextBlock));
            numDisplayFactory.SetBinding(TextBlock.TextProperty,
                new Binding("NewNumber"));
            numDisplayFactory.SetValue(TextBlock.PaddingProperty,
                new Thickness(6, 4, 6, 4));
            numDisplayFactory.SetValue(TextBlock.VerticalAlignmentProperty,
                VerticalAlignment.Center);
            newNumCol.CellTemplate = new DataTemplate
            { VisualTree = numDisplayFactory };

            var numEditFactory = new FrameworkElementFactory(typeof(TextBox));
            numEditFactory.SetBinding(TextBox.TextProperty,
                new Binding("NewNumber")
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                });
            numEditFactory.SetValue(TextBox.BorderThicknessProperty,
                new Thickness(0));
            numEditFactory.SetValue(TextBox.PaddingProperty,
                new Thickness(4, 2, 4, 2));
            numEditFactory.SetValue(TextBox.FontSizeProperty, 12.5);
            newNumCol.CellEditingTemplate = new DataTemplate
            { VisualTree = numEditFactory };

            // Number conflict style
            var numCellStyle = new Style(typeof(DataGridCell));
            var numConflict = new DataTrigger
            {
                Binding = new Binding("HasNumberConflict"),
                Value = true
            };
            numConflict.Setters.Add(new Setter(
                DataGridCell.BackgroundProperty,
                new SolidColorBrush(Color.FromArgb(55, 220, 50, 50))));
            numConflict.Setters.Add(new Setter(
                DataGridCell.ToolTipProperty,
                "Duplicate sheet number \u2013 will be skipped"));
            numCellStyle.Triggers.Add(numConflict);

            var numChanged = new DataTrigger
            {
                Binding = new Binding("NumberChanged"),
                Value = true
            };
            numChanged.Setters.Add(new Setter(
                DataGridCell.ForegroundProperty,
                new SolidColorBrush(Color.FromRgb(0, 100, 180))));
            numCellStyle.Triggers.Add(numChanged);

            newNumCol.CellStyle = numCellStyle;
            dg.Columns.Add(newNumCol);

            // ── Original Name (read-only) ──────────────────────
            var origNameCol = new DataGridTextColumn
            {
                Header = "Original Name",
                Binding = new Binding("OriginalName"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                IsReadOnly = true
            };
            origNameCol.CellStyle = mutedStyle;
            dg.Columns.Add(origNameCol);

            // ── New Name (editable) ────────────────────────────
            var newNameCol = new DataGridTemplateColumn
            {
                Header = "New Name",
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                SortMemberPath = "NewName"
            };

            var nameDispFactory = new FrameworkElementFactory(typeof(TextBlock));
            nameDispFactory.SetBinding(TextBlock.TextProperty,
                new Binding("NewName"));
            nameDispFactory.SetValue(TextBlock.PaddingProperty,
                new Thickness(6, 4, 6, 4));
            nameDispFactory.SetValue(TextBlock.VerticalAlignmentProperty,
                VerticalAlignment.Center);
            newNameCol.CellTemplate = new DataTemplate
            { VisualTree = nameDispFactory };

            var nameEditFactory = new FrameworkElementFactory(typeof(TextBox));
            nameEditFactory.SetBinding(TextBox.TextProperty,
                new Binding("NewName")
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                });
            nameEditFactory.SetValue(TextBox.BorderThicknessProperty,
                new Thickness(0));
            nameEditFactory.SetValue(TextBox.PaddingProperty,
                new Thickness(4, 2, 4, 2));
            nameEditFactory.SetValue(TextBox.FontSizeProperty, 12.5);
            newNameCol.CellEditingTemplate = new DataTemplate
            { VisualTree = nameEditFactory };

            // Changed name style (blue text, no conflict for names)
            var nameCellStyle = new Style(typeof(DataGridCell));
            var nameChangedTrigger = new DataTrigger
            {
                Binding = new Binding("NameChanged"),
                Value = true
            };
            nameChangedTrigger.Setters.Add(new Setter(
                DataGridCell.ForegroundProperty,
                new SolidColorBrush(Color.FromRgb(0, 100, 180))));
            nameCellStyle.Triggers.Add(nameChangedTrigger);
            newNameCol.CellStyle = nameCellStyle;
            dg.Columns.Add(newNameCol);

            // ── Views count (read-only) ────────────────────────
            dg.Columns.Add(new DataGridTextColumn
            {
                Header = "Views",
                Binding = new Binding("ViewCount"),
                Width = new DataGridLength(55),
                IsReadOnly = true
            });

            // Row style: subtle red on number conflict
            var rowStyle = new Style(typeof(DataGridRow));
            var rowConflict = new DataTrigger
            {
                Binding = new Binding("HasNumberConflict"),
                Value = true
            };
            rowConflict.Setters.Add(new Setter(
                DataGridRow.BackgroundProperty,
                new SolidColorBrush(Color.FromArgb(25, 220, 50, 50))));
            rowStyle.Triggers.Add(rowConflict);
            dg.RowStyle = rowStyle;

            dg.CellEditEnding += DataGrid_CellEditEnding;

            return dg;
        }

        // ── Tool row ───────────────────────────────────────────

        private StackPanel BuildToolRow()
        {
            var stack = new StackPanel
            {
                Orientation = Orientation.Vertical,
                Margin = new Thickness(0, 0, 0, 12)
            };

            // --- Number prefix row ---
            var numPrefixRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 6)
            };
            numPrefixRow.Children.Add(Label("# Prefix:", 72));
            var numPrefBorder = MakeInputBorder(180);
            prefixNumBox = MakeInputTextBox();
            numPrefBorder.Child = prefixNumBox;
            numPrefixRow.Children.Add(numPrefBorder);

            var addNumPrefBtn = CreateButton("Add Prefix #", BluePrimary,
                Color.FromRgb(255, 255, 255));
            addNumPrefBtn.Width = 120;
            addNumPrefBtn.Margin = new Thickness(8, 0, 0, 0);
            addNumPrefBtn.Click += AddPrefixNumber_Click;
            numPrefixRow.Children.Add(addNumPrefBtn);

            // Number cut
            numPrefixRow.Children.Add(Spacer(16));
            numPrefixRow.Children.Add(Label("# Cut:", 50));
            var cutNumBorder = MakeInputBorder(140);
            cutNumBox = MakeInputTextBox();
            cutNumBorder.Child = cutNumBox;
            numPrefixRow.Children.Add(cutNumBorder);

            var cutNumBtn = CreateButton("Cut #", OrangeAccent,
                Color.FromRgb(255, 255, 255));
            cutNumBtn.Width = 80;
            cutNumBtn.Margin = new Thickness(8, 0, 0, 0);
            cutNumBtn.Click += CutNumber_Click;
            numPrefixRow.Children.Add(cutNumBtn);

            stack.Children.Add(numPrefixRow);

            // --- Name prefix row ---
            var namePrefixRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 6)
            };
            namePrefixRow.Children.Add(Label("Name Prefix:", 72));
            var namePrefBorder = MakeInputBorder(180);
            prefixNameBox = MakeInputTextBox();
            namePrefBorder.Child = prefixNameBox;
            namePrefixRow.Children.Add(namePrefBorder);

            var addNamePrefBtn = CreateButton("Add Prefix Name", BluePrimary,
                Color.FromRgb(255, 255, 255));
            addNamePrefBtn.Width = 120;
            addNamePrefBtn.Margin = new Thickness(8, 0, 0, 0);
            addNamePrefBtn.Click += AddPrefixName_Click;
            namePrefixRow.Children.Add(addNamePrefBtn);

            // Name cut
            namePrefixRow.Children.Add(Spacer(16));
            namePrefixRow.Children.Add(Label("Name Cut:", 50));
            var cutNameBorder = MakeInputBorder(140);
            cutNameBox = MakeInputTextBox();
            cutNameBorder.Child = cutNameBox;
            namePrefixRow.Children.Add(cutNameBorder);

            var cutNameBtn = CreateButton("Cut Name", OrangeAccent,
                Color.FromRgb(255, 255, 255));
            cutNameBtn.Width = 80;
            cutNameBtn.Margin = new Thickness(8, 0, 0, 0);
            cutNameBtn.Click += CutName_Click;
            namePrefixRow.Children.Add(cutNameBtn);

            stack.Children.Add(namePrefixRow);

            // --- Reset row ---
            var resetRow = new StackPanel
            {
                Orientation = Orientation.Horizontal
            };
            var resetBtn = CreateButton("Reset All", GrayBg,
                Color.FromRgb(80, 80, 80));
            resetBtn.Width = 110;
            resetBtn.Click += ResetAll_Click;
            resetRow.Children.Add(resetBtn);
            stack.Children.Add(resetRow);

            return stack;
        }

        // ── Button row ─────────────────────────────────────────

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

        // ── Events ─────────────────────────────────────────────

        private void Entry_PropertyChanged(object sender,
            PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "NewNumber"
                || e.PropertyName == "NewName")
            {
                CheckConflicts();
                UpdateSummary();
            }
        }

        private void DataGrid_CellEditEnding(object sender,
            DataGridCellEditEndingEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                CheckConflicts();
                UpdateSummary();
            }));
        }

        private void SelectAll_Changed(object sender, RoutedEventArgs e)
        {
            if (_suppressCheckSync) return;
            bool check = selectAllCheckBox.IsChecked == true;
            foreach (var entry in allEntries)
                entry.IsSelected = check;
            UpdateSummary();
        }

        private void RowCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (_suppressCheckSync) return;
            _suppressCheckSync = true;

            var chk = sender as CheckBox;
            bool newState = chk?.IsChecked == true;
            var clicked = chk?.DataContext as SheetAuditEntry;

            var highlighted = dataGrid.SelectedItems
                .OfType<SheetAuditEntry>().ToList();

            if (highlighted.Count > 1
                && clicked != null
                && highlighted.Contains(clicked))
            {
                foreach (var entry in highlighted)
                    entry.IsSelected = newState;
            }

            if (allEntries.All(x => x.IsSelected))
                selectAllCheckBox.IsChecked = true;
            else if (allEntries.All(x => !x.IsSelected))
                selectAllCheckBox.IsChecked = false;
            else
                selectAllCheckBox.IsChecked = null;

            UpdateSummary();
            _suppressCheckSync = false;
        }

        private void SearchBox_TextChanged(object sender,
            TextChangedEventArgs e)
        {
            string filter = searchBox.Text.Trim().ToLower();
            collectionView.Filter = obj =>
            {
                if (string.IsNullOrEmpty(filter)) return true;
                var entry = obj as SheetAuditEntry;
                if (entry == null) return false;
                return entry.OriginalNumber.ToLower().Contains(filter)
                    || entry.OriginalName.ToLower().Contains(filter)
                    || entry.NewNumber.ToLower().Contains(filter)
                    || entry.NewName.ToLower().Contains(filter);
            };
        }

        // ── Prefix / Cut handlers ──────────────────────────────

        private void AddPrefixNumber_Click(object sender, RoutedEventArgs e)
        {
            string prefix = prefixNumBox.Text;
            if (string.IsNullOrEmpty(prefix))
            { ShowMsg("Enter a prefix for the sheet number."); return; }

            int count = 0;
            foreach (var entry in allEntries)
            {
                if (!entry.IsSelected) continue;
                if (!entry.NewNumber.StartsWith(prefix))
                { entry.NewNumber = prefix + entry.NewNumber; count++; }
            }
            Refresh();
            if (count == 0) ShowMsg("No sheets were modified.");
        }

        private void CutNumber_Click(object sender, RoutedEventArgs e)
        {
            string text = cutNumBox.Text;
            if (string.IsNullOrEmpty(text))
            { ShowMsg("Enter the text to remove from sheet numbers."); return; }

            int count = 0;
            foreach (var entry in allEntries)
            {
                if (!entry.IsSelected) continue;
                if (entry.NewNumber.Contains(text))
                { entry.NewNumber = entry.NewNumber.Replace(text, ""); count++; }
            }
            Refresh();
            if (count == 0) ShowMsg("No selected sheets contained the specified text in their number.");
        }

        private void AddPrefixName_Click(object sender, RoutedEventArgs e)
        {
            string prefix = prefixNameBox.Text;
            if (string.IsNullOrEmpty(prefix))
            { ShowMsg("Enter a prefix for the sheet name."); return; }

            int count = 0;
            foreach (var entry in allEntries)
            {
                if (!entry.IsSelected) continue;
                if (!entry.NewName.StartsWith(prefix))
                { entry.NewName = prefix + entry.NewName; count++; }
            }
            Refresh();
            if (count == 0) ShowMsg("No sheets were modified.");
        }

        private void CutName_Click(object sender, RoutedEventArgs e)
        {
            string text = cutNameBox.Text;
            if (string.IsNullOrEmpty(text))
            { ShowMsg("Enter the text to remove from sheet names."); return; }

            int count = 0;
            foreach (var entry in allEntries)
            {
                if (!entry.IsSelected) continue;
                if (entry.NewName.Contains(text))
                { entry.NewName = entry.NewName.Replace(text, ""); count++; }
            }
            Refresh();
            if (count == 0) ShowMsg("No selected sheets contained the specified text in their name.");
        }

        private void ResetAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var entry in allEntries)
            {
                entry.NewNumber = entry.OriginalNumber;
                entry.NewName = entry.OriginalName;
            }
            Refresh();
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            CheckConflicts();

            int changes = allEntries.Count(
                x => x.AnyChanged && !x.HasNumberConflict);
            int conflicts = allEntries.Count(x => x.HasNumberConflict);

            if (changes == 0 && conflicts == 0)
            { ShowMsg("No changes to apply."); return; }

            string msg = $"{changes} sheet(s) will be updated.";
            if (conflicts > 0)
                msg += $"\n{conflicts} sheet(s) have duplicate numbers "
                     + "and will be SKIPPED.";
            msg += "\n\nProceed?";

            var result = MessageBox.Show(msg, "HMV Tools \u2013 Confirm",
                MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes) return;

            Results = allEntries.ToList();
            DialogResult = true;
            Close();
        }

        // ── Conflict detection (sheet numbers only) ────────────

        private void CheckConflicts()
        {
            // Duplicate NewNumber within the table
            var numCounts = new Dictionary<string, int>(
                StringComparer.Ordinal);
            foreach (var entry in allEntries)
            {
                if (!numCounts.ContainsKey(entry.NewNumber))
                    numCounts[entry.NewNumber] = 0;
                numCounts[entry.NewNumber]++;
            }

            foreach (var entry in allEntries)
            {
                bool conflict = false;

                if (numCounts.ContainsKey(entry.NewNumber)
                    && numCounts[entry.NewNumber] > 1)
                    conflict = true;

                // Conflict with external sheets
                if (!conflict
                    && entry.NewNumber != entry.OriginalNumber
                    && externalSheetNumbers.Contains(entry.NewNumber))
                    conflict = true;

                entry.HasNumberConflict = conflict;
            }
        }

        // ── Summary ────────────────────────────────────────────

        private void UpdateSummary()
        {
            int total = allEntries.Count;
            int selected = allEntries.Count(e => e.IsSelected);
            int modified = allEntries.Count(e => e.AnyChanged);
            int conflicts = allEntries.Count(e => e.HasNumberConflict);

            summaryText.Text =
                $"Total: {total}    Selected: {selected}    "
                + $"Modified: {modified}    # Conflicts: {conflicts}";
        }

        // ── Helpers ────────────────────────────────────────────

        private void Refresh()
        {
            CheckConflicts();
            UpdateSummary();
        }

        private void ShowMsg(string text) =>
            MessageBox.Show(text, "HMV Tools", MessageBoxButton.OK);

        private TextBlock Label(string text, double width) =>
            new TextBlock
            {
                Text = text,
                FontSize = 13,
                Width = width,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 0)
            };

        private FrameworkElement Spacer(double w) =>
            new Border { Width = w };

        private Border MakeInputBorder(double width) =>
            new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Width = width
            };

        private TextBox MakeInputTextBox() =>
            new TextBox
            {
                Height = 32,
                FontSize = 13,
                VerticalContentAlignment = VerticalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(10, 0, 10, 0)
            };

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