using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    // ── Plain data classes ──────────────────────────────────────

    public class LinkInfo
    {
        public string Name { get; set; }
        public int Index { get; set; }
    }

    public class SpotElevationSettings
    {
        public int FoundationSourceIndex { get; set; }
        public int FloorLinkIndex { get; set; }
        public double LeaderOffsetMm { get; set; }
        public bool OffsetX { get; set; }
        public bool OffsetY { get; set; }
        public bool UseHmvStandard { get; set; }
        public bool CreateGrid { get; set; }
        public int View3DIndex { get; set; } // Added for 3D View selection
    }

    // ── Window ─────────────────────────────────────────────────

    public class SpotElevationWindow : Window
    {
        private ComboBox cmbFoundationSource;
        private ComboBox cmbFloorLink;
        private ComboBox cmbView3D; // Added 3D View Combo
        private TextBox txtLeaderOffset;
        private CheckBox chkOffsetX;
        private CheckBox chkOffsetY;
        private CheckBox chkGrid;
        private CheckBox chkHmvStandard;

        private List<LinkInfo> linkInfos;
        private List<string> foundationSourceItems;
        private List<string> floorLinkItems;
        private List<string> view3DItems; // Added to store View Names

        public SpotElevationSettings Settings { get; private set; }

        // Colors
        private static readonly Color BluePrimary = Color.FromRgb(0, 120, 212);
        private static readonly Color GrayBg = Color.FromRgb(240, 240, 243);
        private static readonly Color DarkText = Color.FromRgb(30, 30, 30);
        private static readonly Color MutedText = Color.FromRgb(120, 120, 130);
        private static readonly Color BorderColor = Color.FromRgb(200, 200, 210);
        private static readonly Color WindowBg = Color.FromRgb(245, 245, 248);
        private static readonly Color AccentBg = Color.FromRgb(232, 243, 255);

        // Updated constructor to accept view names
        public SpotElevationWindow(List<LinkInfo> links, List<string> viewNames)
        {
            linkInfos = links;
            view3DItems = viewNames;

            Title = "HMV Tools – Spot Elevation on Floor";
            Width = 500;
            Height = 650; // Increased height to fit the new search block
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(WindowBg);

            var main = new Grid { Margin = new Thickness(24) };
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 0 Title
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 1 Foundation
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 2 Floor link
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 3 3D View (NEW)
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 4 Offset
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 5 HMV
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 6 Info
            main.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // 7 Spacer
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 8 Buttons

            // ── Row 0: Title ───────────────────────────────────
            var title = new TextBlock
            {
                Text = "Spot Elevation on Linked Floor",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 16)
            };
            Grid.SetRow(title, 0);
            main.Children.Add(title);

            // ── Row 1: Foundation source (searchable) ──────────
            foundationSourceItems = BuildFoundationSourceItems();
            var foundPanel = CreateSearchableCombo(
                "Foundation source:",
                foundationSourceItems, 0,
                out cmbFoundationSource);
            Grid.SetRow(foundPanel, 1);
            main.Children.Add(foundPanel);

            // ── Row 2: Floor link (searchable) ─────────────────
            floorLinkItems = links.Select(l => l.Name).ToList();
            var floorPanel = CreateSearchableCombo(
                "Floor link (where floors live):",
                floorLinkItems,
                links.Count > 0 ? 0 : -1,
                out cmbFloorLink);
            Grid.SetRow(floorPanel, 2);
            main.Children.Add(floorPanel);

            // ── Row 3: 3D View Selection (searchable) ──────────
            var viewPanel = CreateSearchableCombo(
                "3D View context for raycasting:",
                view3DItems,
                view3DItems.Count > 0 ? 0 : -1,
                out cmbView3D);
            Grid.SetRow(viewPanel, 3);
            main.Children.Add(viewPanel);

            // ── Row 4: Offset controls ─────────────────────────
            var offsetRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 12)
            };
            offsetRow.Children.Add(new TextBlock
            {
                Text = "Leader offset (mm):",
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 10, 0)
            });

            var offsetBorder = CreateInputBorder(100);
            txtLeaderOffset = new TextBox
            {
                Text = "500",
                Height = 32,
                FontSize = 13,
                VerticalContentAlignment = VerticalAlignment.Center,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(8, 0, 8, 0)
            };
            txtLeaderOffset.PreviewTextInput += NumericOnly;
            WireFocus(txtLeaderOffset, offsetBorder);
            offsetBorder.Child = txtLeaderOffset;
            offsetRow.Children.Add(offsetBorder);

            chkOffsetX = MakeCheckBox("X", true, 16);
            offsetRow.Children.Add(chkOffsetX);

            chkOffsetY = MakeCheckBox("Y", true, 8);
            offsetRow.Children.Add(chkOffsetY);

            chkGrid = MakeCheckBox("Grid?", false, 16);
            offsetRow.Children.Add(chkGrid);

            Grid.SetRow(offsetRow, 4);
            main.Children.Add(offsetRow);

            // ── Row 5: HMV Standard ────────────────────────────
            var hmvBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BluePrimary),
                BorderThickness = new Thickness(1),
                Background = new SolidColorBrush(AccentBg),
                Padding = new Thickness(12, 10, 12, 10),
                Margin = new Thickness(0, 0, 0, 12)
            };
            var hmvPanel = new StackPanel();
            chkHmvStandard = new CheckBox
            {
                Content = "HMV Standard (NTCE / NAP)",
                FontSize = 13,
                FontWeight = FontWeights.Medium,
                Foreground = new SolidColorBrush(DarkText),
                IsChecked = false
            };
            hmvPanel.Children.Add(chkHmvStandard);
            hmvPanel.Children.Add(new TextBlock
            {
                Text = "Places two spot elevations per element:\n"
                     + "  • N.T.C.E. → top of selected element\n"
                     + "  • N.A.P.   → intersection with linked floor",
                FontSize = 11,
                Foreground = new SolidColorBrush(MutedText),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(20, 4, 0, 0)
            });
            hmvBorder.Child = hmvPanel;
            Grid.SetRow(hmvBorder, 5);
            main.Children.Add(hmvBorder);

            // ── Row 6: Info ────────────────────────────────────
            var info = new TextBlock
            {
                Text = "After clicking OK you will pick elements in the canvas.\n"
                     + "A vertical ray finds the floor face in the chosen link.",
                FontSize = 11,
                Foreground = new SolidColorBrush(MutedText),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 4, 0, 0)
            };
            Grid.SetRow(info, 6);
            main.Children.Add(info);

            // ── Row 8: Buttons ─────────────────────────────────
            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var cancelBtn = CreateButton("Cancel", GrayBg, Color.FromRgb(60, 60, 60));
            cancelBtn.Width = 90;
            cancelBtn.Margin = new Thickness(0, 0, 8, 0);
            cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };

            var okBtn = CreateButton("OK – Pick Elements", BluePrimary, Colors.White);
            okBtn.Width = 180;
            okBtn.Click += (s, e) => Accept();

            btnPanel.Children.Add(cancelBtn);
            btnPanel.Children.Add(okBtn);
            Grid.SetRow(btnPanel, 8);
            main.Children.Add(btnPanel);

            Content = main;
        }

        // ── Build items ────────────────────────────────────────

        private List<string> BuildFoundationSourceItems()
        {
            var items = new List<string> { "Active Model" };
            items.AddRange(linkInfos.Select(l => l.Name));
            return items;
        }

        // ── Accept ─────────────────────────────────────────────

        private void Accept()
        {
            if (cmbFloorLink.SelectedIndex < 0)
            {
                MessageBox.Show("Select the link that contains the floors.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            if (!double.TryParse(txtLeaderOffset.Text, out double offset)
                || offset <= 0)
            {
                MessageBox.Show("Enter a valid leader offset > 0.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            // Resolve indices from potentially filtered combo items
            string foundText = cmbFoundationSource.SelectedItem as string;
            int foundSrcIdx = foundText != null
                ? foundationSourceItems.IndexOf(foundText) - 1
                : -1;

            string floorText = cmbFloorLink.SelectedItem as string;
            int floorLinkIdx = floorText != null
                ? floorLinkItems.IndexOf(floorText)
                : -1;

            if (floorLinkIdx < 0)
            {
                MessageBox.Show("Select a valid floor link.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            string viewText = cmbView3D.SelectedItem as string;
            int viewIdx = viewText != null
                ? view3DItems.IndexOf(viewText)
                : -1;

            if (viewIdx < 0)
            {
                MessageBox.Show("Select a valid 3D View.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            Settings = new SpotElevationSettings
            {
                FoundationSourceIndex = foundSrcIdx,
                FloorLinkIndex = floorLinkIdx,
                LeaderOffsetMm = offset,
                OffsetX = chkOffsetX.IsChecked == true,
                OffsetY = chkOffsetY.IsChecked == true,
                UseHmvStandard = chkHmvStandard.IsChecked == true,
                CreateGrid = chkGrid.IsChecked == true,
                View3DIndex = viewIdx // Save the index!
            };

            DialogResult = true;
            Close();
        }

        // ── Searchable combo ───────────────────────────────────

        private StackPanel CreateSearchableCombo(
            string label, List<string> allItems,
            int selectedIndex, out ComboBox combo)
        {
            var panel = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };
            panel.Children.Add(new TextBlock
            {
                Text = label,
                FontSize = 13,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 4)
            });

            // Search box
            var searchBorder = CreateInputBorder(0);
            searchBorder.Margin = new Thickness(0, 0, 0, 4);
            var searchBox = new TextBox
            {
                Height = 28,
                FontSize = 12,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(8, 0, 8, 0),
                VerticalContentAlignment = VerticalAlignment.Center,
                Foreground = new SolidColorBrush(MutedText),
                Text = "Search..."
            };
            searchBox.GotFocus += (s, e) =>
            {
                if (searchBox.Text == "Search...")
                {
                    searchBox.Text = "";
                    searchBox.Foreground = new SolidColorBrush(DarkText);
                }
                searchBorder.BorderBrush = new SolidColorBrush(BluePrimary);
            };
            searchBox.LostFocus += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(searchBox.Text))
                {
                    searchBox.Text = "Search...";
                    searchBox.Foreground = new SolidColorBrush(MutedText);
                }
                searchBorder.BorderBrush = new SolidColorBrush(BorderColor);
            };
            searchBorder.Child = searchBox;
            panel.Children.Add(searchBorder);

            // Combo
            var comboBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White
            };
            combo = new ComboBox
            {
                Height = 34,
                FontSize = 13,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(8, 0, 8, 0),
                VerticalContentAlignment = VerticalAlignment.Center
            };
            foreach (var item in allItems) combo.Items.Add(item);
            if (selectedIndex >= 0 && selectedIndex < allItems.Count)
                combo.SelectedIndex = selectedIndex;

            ComboBox capturedCombo = combo;
            searchBox.TextChanged += (s, e) =>
            {
                string filter = searchBox.Text;
                if (filter == "Search...") filter = "";
                FilterCombo(capturedCombo, allItems, filter);
            };

            comboBorder.Child = combo;
            panel.Children.Add(comboBorder);
            return panel;
        }

        private void FilterCombo(ComboBox combo, List<string> all, string filter)
        {
            string selected = combo.SelectedItem as string;
            combo.Items.Clear();
            var filtered = string.IsNullOrWhiteSpace(filter)
                ? all
                : all.Where(i => i.IndexOf(filter,
                    StringComparison.OrdinalIgnoreCase) >= 0).ToList();
            foreach (var item in filtered) combo.Items.Add(item);
            if (selected != null && filtered.Contains(selected))
                combo.SelectedItem = selected;
            else if (filtered.Count == 1)
                combo.SelectedIndex = 0;
        }

        // ── UI helpers ─────────────────────────────────────────

        private Border CreateInputBorder(double width)
        {
            var b = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White
            };
            if (width > 0) b.Width = width;
            return b;
        }

        private void WireFocus(TextBox tb, Border b)
        {
            tb.GotFocus += (s, e) =>
                b.BorderBrush = new SolidColorBrush(BluePrimary);
            tb.LostFocus += (s, e) =>
                b.BorderBrush = new SolidColorBrush(BorderColor);
        }

        private CheckBox MakeCheckBox(string text, bool isChecked, double leftMargin)
        {
            return new CheckBox
            {
                Content = text,
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(leftMargin, 0, 0, 0),
                IsChecked = isChecked
            };
        }

        private void NumericOnly(object sender, TextCompositionEventArgs e)
        {
            foreach (char c in e.Text)
            {
                if (!char.IsDigit(c) && c != '.') { e.Handled = true; return; }
                if (c == '.' && txtLeaderOffset.Text.Contains("."))
                { e.Handled = true; return; }
            }
        }

        private Button CreateButton(string text, Color bgColor, Color fgColor)
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
            var tmpl = new ControlTemplate(typeof(Button));
            var bd = new FrameworkElementFactory(typeof(Border));
            bd.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
            bd.SetValue(Border.BackgroundProperty, new SolidColorBrush(bgColor));
            bd.SetValue(Border.PaddingProperty, new Thickness(14, 6, 14, 6));
            var cp = new FrameworkElementFactory(typeof(ContentPresenter));
            cp.SetValue(ContentPresenter.HorizontalAlignmentProperty,
                HorizontalAlignment.Center);
            cp.SetValue(ContentPresenter.VerticalAlignmentProperty,
                VerticalAlignment.Center);
            bd.AppendChild(cp);
            tmpl.VisualTree = bd;
            btn.Template = tmpl;
            return btn;
        }
    }
}