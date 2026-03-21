using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    // ── Plain data classes (no Revit references) ───────────────

    public class LinkInfo
    {
        public string Name { get; set; }
        public int Index { get; set; }
    }

    public class SpotElevationSettings
    {
        /// <summary>-1 = Active Model, otherwise index into LinkInfos list.</summary>
        public int FoundationSourceIndex { get; set; }
        public int FloorLinkIndex { get; set; }
        public double LeaderOffsetMm { get; set; }
        public bool OffsetX { get; set; }
        public bool OffsetY { get; set; }
        public bool UseHmvStandard { get; set; }
        public bool CreateGrid { get; set; }
    }

    // ── Window ─────────────────────────────────────────────────

    public class SpotElevationWindow : Window
    {
        // Controls
        private ComboBox cmbFoundationSource;
        private ComboBox cmbFloorLink;
        private TextBox txtLeaderOffset;
        private CheckBox chkOffsetX;
        private CheckBox chkOffsetY;
        private CheckBox chkGrid;
        private CheckBox chkHmvStandard;

        // Data
        private List<LinkInfo> linkInfos;
        private List<string> foundationSourceItems;
        private List<string> floorLinkItems;

        /// <summary>User's settings, or null if cancelled.</summary>
        public SpotElevationSettings Settings { get; private set; }

        // Colors (same palette as PipeAnnotationWindow)
        private static readonly Color BluePrimary = Color.FromRgb(0, 120, 212);
        private static readonly Color GrayBg = Color.FromRgb(240, 240, 243);
        private static readonly Color DarkText = Color.FromRgb(30, 30, 30);
        private static readonly Color MutedText = Color.FromRgb(120, 120, 130);
        private static readonly Color BorderColor = Color.FromRgb(200, 200, 210);
        private static readonly Color WindowBg = Color.FromRgb(245, 245, 248);
        private static readonly Color AccentBg = Color.FromRgb(232, 243, 255);
        private static readonly Color AccentBorder = Color.FromRgb(0, 120, 212);

        public SpotElevationWindow(List<LinkInfo> links)
        {
            linkInfos = links;

            Title = "HMV Tools – Spot Elevation on Floor";
            Width = 500;
            Height = 550;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(WindowBg);

            var main = new Grid { Margin = new Thickness(24) };
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                        // 0 Title
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                        // 1 Foundation source
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                        // 2 Floor link
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                        // 3 Leader offset + both axis
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                        // 4 HMV Standard checkbox
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                        // 5 Info
            main.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });   // 6 Spacer
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                        // 7 Buttons

            // ── Row 0: Title ───────────────────────────────────────
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

            // ── Row 1: Foundation source ───────────────────────────
            foundationSourceItems = BuildFoundationSourceItems();
            var foundPanel = CreateSearchableCombo(
                "Foundation source:",
                foundationSourceItems,
                0,
                out cmbFoundationSource);
            Grid.SetRow(foundPanel, 1);
            main.Children.Add(foundPanel);

            // ── Row 2: Floor link ──────────────────────────────────
            floorLinkItems = links.Select(l => l.Name).ToList();
            var floorPanel = CreateSearchableCombo(
                "Floor link (where floors live):",
                floorLinkItems,
                links.Count > 0 ? 0 : -1,
                out cmbFloorLink);
            Grid.SetRow(floorPanel, 2);
            main.Children.Add(floorPanel);

            // ── Row 3: Leader offset + Both axis ───────────────────
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

            var offsetBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Width = 100
            };
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
            txtLeaderOffset.PreviewTextInput += (s, e) =>
            {
                e.Handled = !IsNumericInput(e.Text, txtLeaderOffset.Text);
            };
            txtLeaderOffset.GotFocus += (s, e) =>
                offsetBorder.BorderBrush = new SolidColorBrush(BluePrimary);
            txtLeaderOffset.LostFocus += (s, e) =>
                offsetBorder.BorderBrush = new SolidColorBrush(BorderColor);
            offsetBorder.Child = txtLeaderOffset;
            offsetRow.Children.Add(offsetBorder);

            chkOffsetX = new CheckBox
            {
                Content = "X",
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(16, 0, 0, 0),
                IsChecked = true
            };
            offsetRow.Children.Add(chkOffsetX);

            chkOffsetY = new CheckBox
            {
                Content = "Y",
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(8, 0, 0, 0),
                IsChecked = true
            };
            offsetRow.Children.Add(chkOffsetY);

            chkGrid = new CheckBox
            {
                Content = "Grid?",
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(16, 0, 0, 0),
                IsChecked = false
            };
            offsetRow.Children.Add(chkGrid);

            Grid.SetRow(offsetRow, 3);
            main.Children.Add(offsetRow);

            // ── Row 4: HMV Standard checkbox ───────────────────────
            var hmvBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(AccentBorder),
                BorderThickness = new Thickness(1),
                Background = new SolidColorBrush(AccentBg),
                Padding = new Thickness(12, 10, 12, 10),
                Margin = new Thickness(0, 0, 0, 12)
            };
            var hmvPanel = new StackPanel();
            chkHmvStandard = new CheckBox
            {
                FontSize = 13,
                FontWeight = FontWeights.Medium,
                Foreground = new SolidColorBrush(DarkText),
                IsChecked = false
            };
            chkHmvStandard.Content = "HMV Standard (NTCE / NAP)";

            var hmvDesc = new TextBlock
            {
                Text = "Places two spot elevations per foundation:\n"
                     + "  • N.T.C.E. → top of the selected element\n"
                     + "  • N.A.P.   → intersection with the linked floor",
                FontSize = 11,
                Foreground = new SolidColorBrush(MutedText),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(20, 4, 0, 0)
            };
            hmvPanel.Children.Add(chkHmvStandard);
            hmvPanel.Children.Add(hmvDesc);
            hmvBorder.Child = hmvPanel;

            Grid.SetRow(hmvBorder, 4);
            main.Children.Add(hmvBorder);

            // ── Row 5: Info text ───────────────────────────────────
            var info = new TextBlock
            {
                Text = "After clicking OK you will pick elements in the canvas.\n"
                     + "A vertical ray finds the floor face in the chosen link.",
                FontSize = 11,
                Foreground = new SolidColorBrush(MutedText),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 4, 0, 0)
            };
            Grid.SetRow(info, 5);
            main.Children.Add(info);

            // ── Row 7: Buttons ─────────────────────────────────────
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
            Grid.SetRow(btnPanel, 7);
            main.Children.Add(btnPanel);

            Content = main;
        }

        // ── Build combo items ──────────────────────────────────────

        private List<string> BuildFoundationSourceItems()
        {
            var items = new List<string> { "Active Model" };
            items.AddRange(linkInfos.Select(l => l.Name));
            return items;
        }

        // ── Accept ─────────────────────────────────────────────────

        private void Accept()
        {
            if (cmbFloorLink.SelectedIndex < 0)
            {
                MessageBox.Show("Select the link that contains the floors.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            if (!double.TryParse(txtLeaderOffset.Text, out double offset) || offset <= 0)
            {
                MessageBox.Show("Enter a valid leader offset greater than 0.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            string foundText = cmbFoundationSource.SelectedItem as string;
            int foundSrcIdx = foundText != null
                ? foundationSourceItems.IndexOf(foundText) - 1
                : -1;

            string floorText = cmbFloorLink.SelectedItem as string;
            int floorLinkIdx = floorText != null
                ? floorLinkItems.IndexOf(floorText)
                : -1;

            Settings = new SpotElevationSettings
            {
                FoundationSourceIndex = foundSrcIdx,
                FloorLinkIndex = floorLinkIdx,
                LeaderOffsetMm = offset,
                OffsetX = chkOffsetX.IsChecked == true,
                OffsetY = chkOffsetY.IsChecked == true,
                UseHmvStandard = chkHmvStandard.IsChecked == true,
                CreateGrid = chkGrid.IsChecked == true
            };

            DialogResult = true;
            Close();
        }

        // ── Helpers ────────────────────────────────────────────────

        private StackPanel CreateSearchableCombo(
            string label,
            List<string> allItems,
            int selectedIndex,
            out ComboBox combo)
        {
            var panel = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };
            panel.Children.Add(new TextBlock
            {
                Text = label,
                FontSize = 13,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 4)
            });

            var searchBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 4)
            };
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
            foreach (var item in allItems)
                combo.Items.Add(item);
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

        private void FilterCombo(ComboBox combo, List<string> allItems, string filter)
        {
            string selected = combo.SelectedItem as string;
            combo.Items.Clear();
            var filtered = string.IsNullOrWhiteSpace(filter)
                ? allItems
                : allItems.Where(i => i.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
            foreach (var item in filtered)
                combo.Items.Add(item);
            if (selected != null && filtered.Contains(selected))
                combo.SelectedItem = selected;
            else if (filtered.Count == 1)
                combo.SelectedIndex = 0;
        }

        private bool IsNumericInput(string newText, string currentText)
        {
            foreach (char c in newText)
            {
                if (!char.IsDigit(c) && c != '.') return false;
                if (c == '.' && currentText.Contains(".")) return false;
            }
            return true;
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
            btn.Template = GetRoundButtonTemplate(bgColor);
            return btn;
        }

        private ControlTemplate GetRoundButtonTemplate(Color bgColor)
        {
            var template = new ControlTemplate(typeof(Button));
            var border = new FrameworkElementFactory(typeof(Border));
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
            border.SetValue(Border.BackgroundProperty, new SolidColorBrush(bgColor));
            border.SetValue(Border.PaddingProperty, new Thickness(14, 6, 14, 6));

            var content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            content.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);

            border.AppendChild(content);
            template.VisualTree = border;
            return template;
        }
    }
}