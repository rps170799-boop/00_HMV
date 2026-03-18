using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    // ── Data container (no Revit references) ───────────────────

    public class GridLevelViewItem
    {
        public int Index { get; set; }
        public string DisplayName { get; set; }
    }

    // ── Window ─────────────────────────────────────────────────

    public class GridLevelExtentWindow : Window
    {
        // Controls
        private RadioButton radio2D;
        private RadioButton radio3D;
        private CheckBox chkGrids;
        private CheckBox chkLevels;
        private ListBox viewList;
        private TextBlock countText;
        private TextBox searchBox;

        // Data
        private List<GridLevelViewItem> allItems;
        private List<bool> checkedState;

        // Colors (matching DwgConvertWindow / PipeAnnotationWindow)
        private static readonly System.Windows.Media.Color BluePrimary =
            System.Windows.Media.Color.FromRgb(0, 120, 212);
        private static readonly System.Windows.Media.Color GrayBg =
            System.Windows.Media.Color.FromRgb(240, 240, 243);
        private static readonly System.Windows.Media.Color DarkText =
            System.Windows.Media.Color.FromRgb(30, 30, 30);
        private static readonly System.Windows.Media.Color MutedText =
            System.Windows.Media.Color.FromRgb(120, 120, 130);
        private static readonly System.Windows.Media.Color BorderColor =
            System.Windows.Media.Color.FromRgb(200, 200, 210);
        private static readonly System.Windows.Media.Color WindowBg =
            System.Windows.Media.Color.FromRgb(245, 245, 248);

        // ── Public results ──
        public bool ConvertTo2D { get; private set; }
        public bool ProcessGrids { get; private set; }
        public bool ProcessLevels { get; private set; }
        public List<int> SelectedIndices { get; private set; }

        public GridLevelExtentWindow(List<GridLevelViewItem> items)
        {
            allItems = items;
            checkedState = new List<bool>(new bool[items.Count]);

            Title = "HMV Tools \u2013 Grid & Level Extent";
            Width = 560;
            Height = 600;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(WindowBg);

            var mainGrid = new System.Windows.Controls.Grid();
            mainGrid.Margin = new Thickness(20);
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // 0 title
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // 1 mode
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // 2 elements
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // 3 views header
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // 4 search
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // 5 list
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // 6 select all
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // 7 info
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // 8 buttons

            // ═══ Row 0 — Title ═══
            var title = new TextBlock
            {
                Text = "Grid & Level Extent",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 12)
            };
            System.Windows.Controls.Grid.SetRow(title, 0);
            mainGrid.Children.Add(title);

            // ═══ Row 1 — Conversion Mode ═══
            var modePanel = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };
            modePanel.Children.Add(new TextBlock
            {
                Text = "Conversion Mode",
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 6)
            });

            var radioPanel = new StackPanel { Orientation = Orientation.Horizontal };

            radio2D = new RadioButton
            {
                Content = "2D \u2014 ViewSpecific",
                GroupName = "ExtentMode",
                IsChecked = true,
                FontSize = 13,
                Cursor = Cursors.Hand,
                Margin = new Thickness(0, 0, 16, 0)
            };
            radioPanel.Children.Add(radio2D);

            radio3D = new RadioButton
            {
                Content = "3D \u2014 Model",
                GroupName = "ExtentMode",
                IsChecked = false,
                FontSize = 13,
                Cursor = Cursors.Hand
            };
            radioPanel.Children.Add(radio3D);

            modePanel.Children.Add(radioPanel);
            System.Windows.Controls.Grid.SetRow(modePanel, 1);
            mainGrid.Children.Add(modePanel);

            // ═══ Row 2 — Elements to Process ═══
            var elemPanel = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };
            elemPanel.Children.Add(new TextBlock
            {
                Text = "Elements to Process",
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 6)
            });

            var chkPanel = new StackPanel { Orientation = Orientation.Horizontal };

            chkGrids = new CheckBox
            {
                Content = "Grids",
                IsChecked = true,
                FontSize = 13,
                Margin = new Thickness(0, 0, 24, 0),
                Cursor = Cursors.Hand
            };
            chkPanel.Children.Add(chkGrids);

            chkLevels = new CheckBox
            {
                Content = "Levels",
                IsChecked = true,
                FontSize = 13,
                Cursor = Cursors.Hand
            };
            chkPanel.Children.Add(chkLevels);

            elemPanel.Children.Add(chkPanel);
            System.Windows.Controls.Grid.SetRow(elemPanel, 2);
            mainGrid.Children.Add(elemPanel);

            // ═══ Row 3 — Views header + count ═══
            var headerPanel = new DockPanel { Margin = new Thickness(0, 0, 0, 4) };
            var viewsLabel = new TextBlock
            {
                Text = "Views",
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center
            };
            headerPanel.Children.Add(viewsLabel);

            countText = new TextBlock
            {
                FontSize = 12,
                Foreground = new SolidColorBrush(MutedText),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Center
            };
            DockPanel.SetDock(countText, Dock.Right);
            headerPanel.Children.Add(countText);
            System.Windows.Controls.Grid.SetRow(headerPanel, 3);
            mainGrid.Children.Add(headerPanel);

            // ═══ Row 4 — Search box ═══
            var searchBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 8)
            };
            searchBox = new TextBox
            {
                Height = 32,
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
            System.Windows.Controls.Grid.SetRow(searchBorder, 4);
            mainGrid.Children.Add(searchBorder);

            // ═══ Row 5 — View list with checkboxes ═══
            var listBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 8)
            };
            viewList = new ListBox
            {
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                FontSize = 13,
                Padding = new Thickness(4)
            };
            listBorder.Child = viewList;
            System.Windows.Controls.Grid.SetRow(listBorder, 5);
            mainGrid.Children.Add(listBorder);

            // ═══ Row 6 — Select All Views button ═══
            var selectAllBtn = CreateButton("Select ALL Views", GrayBg,
                System.Windows.Media.Color.FromRgb(60, 60, 60));
            selectAllBtn.HorizontalAlignment = HorizontalAlignment.Stretch;
            selectAllBtn.Margin = new Thickness(0, 0, 0, 8);
            selectAllBtn.Click += SelectAllViews_Click;
            System.Windows.Controls.Grid.SetRow(selectAllBtn, 6);
            mainGrid.Children.Add(selectAllBtn);

            // ═══ Row 7 — Info text ═══
            var infoText = new TextBlock
            {
                Text = "2D (ViewSpecific): each view controls its own grid/level extents independently.\n"
                     + "3D (Model): moving an extent in one view affects all views.",
                FontSize = 11,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(0, 0, 0, 12),
                TextWrapping = TextWrapping.Wrap
            };
            System.Windows.Controls.Grid.SetRow(infoText, 7);
            mainGrid.Children.Add(infoText);

            // ═══ Row 8 — Action buttons ═══
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var cancelBtn = CreateButton("Cancel", GrayBg,
                System.Windows.Media.Color.FromRgb(60, 60, 60));
            cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };
            cancelBtn.Margin = new Thickness(0, 0, 8, 0);
            cancelBtn.Width = 90;

            var executeBtn = CreateButton("Execute", BluePrimary,
                System.Windows.Media.Color.FromRgb(255, 255, 255));
            executeBtn.Click += Execute_Click;
            executeBtn.Width = 130;

            buttonPanel.Children.Add(cancelBtn);
            buttonPanel.Children.Add(executeBtn);
            System.Windows.Controls.Grid.SetRow(buttonPanel, 8);
            mainGrid.Children.Add(buttonPanel);

            Content = mainGrid;

            PopulateViewList();
            UpdateCount();

            Loaded += (s, e) => searchBox.Focus();
        }

        // ── Populate the ListBox with CheckBox items ──
        private void PopulateViewList()
        {
            viewList.Items.Clear();

            string filter = (searchBox != null)
                ? searchBox.Text.ToLower() : "";

            for (int i = 0; i < allItems.Count; i++)
            {
                if (!string.IsNullOrEmpty(filter)
                    && !allItems[i].DisplayName.ToLower().Contains(filter))
                    continue;

                int idx = i; // capture for closure
                var cb = new CheckBox
                {
                    Content = allItems[i].DisplayName,
                    IsChecked = checkedState[i],
                    FontSize = 13,
                    Padding = new Thickness(4, 4, 4, 4),
                    Cursor = Cursors.Hand
                };
                cb.Checked += (s, e) => { checkedState[idx] = true; UpdateCount(); };
                cb.Unchecked += (s, e) => { checkedState[idx] = false; UpdateCount(); };

                viewList.Items.Add(cb);
            }
        }

        // ── Search filter ──
        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            PopulateViewList();
            UpdateCount();
        }

        // ── Select ALL Views (toggle) ──
        private void SelectAllViews_Click(object sender, RoutedEventArgs e)
        {
            bool allChecked = checkedState.All(c => c);

            for (int i = 0; i < checkedState.Count; i++)
                checkedState[i] = !allChecked;

            PopulateViewList();
            UpdateCount();
        }

        // ── Execute ──
        private void Execute_Click(object sender, RoutedEventArgs e)
        {
            ConvertTo2D = radio2D.IsChecked == true;
            ProcessGrids = chkGrids.IsChecked == true;
            ProcessLevels = chkLevels.IsChecked == true;

            SelectedIndices = new List<int>();
            for (int i = 0; i < allItems.Count; i++)
            {
                if (checkedState[i])
                    SelectedIndices.Add(allItems[i].Index);
            }

            if (!ProcessGrids && !ProcessLevels)
            {
                MessageBox.Show(
                    "Select at least Grids, Levels, or both.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            if (SelectedIndices.Count == 0)
            {
                MessageBox.Show(
                    "Select at least one view.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            DialogResult = true;
            Close();
        }

        // ── Counter ──
        private void UpdateCount()
        {
            int sel = checkedState.Count(c => c);
            countText.Text = $"{sel} / {allItems.Count} selected";
        }

        // ── UI helpers (matching DwgConvertWindow / PipeAnnotationWindow) ──

        private Button CreateButton(string text,
            System.Windows.Media.Color bgColor,
            System.Windows.Media.Color fgColor)
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

        private ControlTemplate GetRoundButtonTemplate(
            System.Windows.Media.Color bgColor)
        {
            var template = new ControlTemplate(typeof(Button));
            var border = new FrameworkElementFactory(typeof(Border));
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
            border.SetValue(Border.BackgroundProperty,
                new SolidColorBrush(bgColor));
            border.SetValue(Border.PaddingProperty,
                new Thickness(14, 6, 14, 6));

            var content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(ContentPresenter.HorizontalAlignmentProperty,
                HorizontalAlignment.Center);
            content.SetValue(ContentPresenter.VerticalAlignmentProperty,
                VerticalAlignment.Center);

            border.AppendChild(content);
            template.VisualTree = border;
            return template;
        }
    }
}