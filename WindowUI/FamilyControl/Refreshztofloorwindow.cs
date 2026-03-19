using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    /// <summary>
    /// Pre-execution window: shows element count, lets user pick
    /// which Revit link contains the target floor.
    /// Pure code-behind WPF — no XAML.
    /// </summary>
    public class RefreshZToFloorWindow : Window
    {
        // Controls
        private ListBox linkListBox;
        private TextBox searchBox;

        // Data
        private List<LinkPickEntry> allLinks;

        // Colors (matching PipeAnnotationWindow palette)
        private static readonly Color BluePrimary = Color.FromRgb(0, 120, 212);
        private static readonly Color GrayBg = Color.FromRgb(240, 240, 243);
        private static readonly Color DarkText = Color.FromRgb(30, 30, 30);
        private static readonly Color MutedText = Color.FromRgb(120, 120, 130);
        private static readonly Color BorderColor = Color.FromRgb(200, 200, 210);
        private static readonly Color WindowBg = Color.FromRgb(245, 245, 248);
        private static readonly Color AccentBg = Color.FromRgb(232, 243, 255);
        private static readonly Color AccentBorder = Color.FromRgb(0, 120, 212);

        /// <summary>ElementId.IntegerValue of the chosen link, or -1.</summary>
        public int SelectedLinkId { get; private set; } = -1;

        public RefreshZToFloorWindow(
            int elementCount,
            List<LinkPickEntry> links)
        {
            allLinks = links;

            Title = "HMV Tools – Refresh Z to Linked Floor";
            Width = 500;
            Height = 480;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(WindowBg);

            var mainGrid = new Grid();
            mainGrid.Margin = new Thickness(20);
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // 0 Title
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // 1 Info badge
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // 2 Subtitle
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // 3 Search
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // 4 List
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // 5 Hint
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // 6 Buttons

            // ── Row 0: Title ───────────────────────────────────────
            var title = new TextBlock
            {
                Text = "Refresh Z to Linked Floor",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 10)
            };
            Grid.SetRow(title, 0);
            mainGrid.Children.Add(title);

            // ── Row 1: Info badge — element count ──────────────────
            var infoBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                Background = new SolidColorBrush(AccentBg),
                BorderBrush = new SolidColorBrush(AccentBorder),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(14, 8, 14, 8),
                Margin = new Thickness(0, 0, 0, 12),
                HorizontalAlignment = HorizontalAlignment.Left
            };
            string plural = elementCount == 1 ? "element" : "elements";
            var infoText = new TextBlock
            {
                Text = $"✔  {elementCount} {plural} selected",
                FontSize = 13,
                Foreground = new SolidColorBrush(AccentBorder)
            };
            infoBorder.Child = infoText;
            Grid.SetRow(infoBorder, 1);
            mainGrid.Children.Add(infoBorder);

            // ── Row 2: Subtitle ────────────────────────────────────
            var subtitle = new TextBlock
            {
                Text = "Select the Revit Link that contains the target floor",
                FontSize = 12,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(0, 0, 0, 6)
            };
            Grid.SetRow(subtitle, 2);
            mainGrid.Children.Add(subtitle);

            // ── Row 3: Search box ──────────────────────────────────
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
            Grid.SetRow(searchBorder, 3);
            mainGrid.Children.Add(searchBorder);

            // ── Row 4: Link list ───────────────────────────────────
            var listBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 10)
            };
            linkListBox = new ListBox
            {
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                FontSize = 13,
                Padding = new Thickness(4),
                SelectionMode = SelectionMode.Single
            };
            linkListBox.MouseDoubleClick += (s, e) =>
            {
                if (linkListBox.SelectedItem != null) Accept();
            };
            listBorder.Child = linkListBox;
            Grid.SetRow(listBorder, 4);
            mainGrid.Children.Add(listBorder);

            // ── Row 5: Hint text ───────────────────────────────────
            var hint = new TextBlock
            {
                Text = "For each element the command shoots a vertical ray\n"
                     + "at (X, Y) and snaps Z to the top of the nearest floor.",
                FontSize = 11,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(0, 0, 0, 12),
                TextWrapping = TextWrapping.Wrap
            };
            Grid.SetRow(hint, 5);
            mainGrid.Children.Add(hint);

            // ── Row 6: Buttons ─────────────────────────────────────
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var cancelBtn = CreateButton("Cancel", GrayBg,
                Color.FromRgb(60, 60, 60));
            cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };
            cancelBtn.Margin = new Thickness(0, 0, 8, 0);
            cancelBtn.Width = 90;

            var okBtn = CreateButton("Refresh Z", BluePrimary,
                Color.FromRgb(255, 255, 255));
            okBtn.Click += (s, e) => Accept();
            okBtn.Width = 140;

            buttonPanel.Children.Add(cancelBtn);
            buttonPanel.Children.Add(okBtn);
            Grid.SetRow(buttonPanel, 6);
            mainGrid.Children.Add(buttonPanel);

            Content = mainGrid;

            // Initial population
            PopulateList(allLinks);

            Loaded += (s, e) => searchBox.Focus();
        }

        // ── Accept ─────────────────────────────────────────────────

        private void Accept()
        {
            if (linkListBox.SelectedItem == null)
            {
                MessageBox.Show("Select a Revit link.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            string selectedText =
                (linkListBox.SelectedItem as TextBlock)?.Text ?? "";
            var entry = allLinks.FirstOrDefault(
                e => e.Name == selectedText);

            if (entry == null)
            {
                MessageBox.Show("Could not resolve selected link.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            SelectedLinkId = entry.LinkId;
            DialogResult = true;
            Close();
        }

        // ── List population ────────────────────────────────────────

        private void PopulateList(List<LinkPickEntry> entries)
        {
            linkListBox.Items.Clear();
            foreach (var e in entries)
                linkListBox.Items.Add(CreateListItem(e.Name));

            if (linkListBox.Items.Count > 0)
                linkListBox.SelectedIndex = 0;
        }

        // ── Search filter ──────────────────────────────────────────

        private void SearchBox_TextChanged(
            object sender, TextChangedEventArgs e)
        {
            linkListBox.Items.Clear();
            string filter = searchBox.Text.ToLower();
            foreach (var entry in allLinks)
            {
                if (entry.Name.ToLower().Contains(filter))
                    linkListBox.Items.Add(CreateListItem(entry.Name));
            }
        }

        // ── UI helpers (matching PipeAnnotationWindow) ─────────────

        private TextBlock CreateListItem(string text)
        {
            return new TextBlock
            {
                Text = text,
                Padding = new Thickness(8, 6, 8, 6)
            };
        }

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
    }

    // ════════════════════════════════════════════════════════════
    //  Report window — same pattern as TextAuditCommand report
    // ════════════════════════════════════════════════════════════

    public class RefreshZReportWindow : Window
    {
        private static readonly Color DarkText = Color.FromRgb(30, 30, 30);
        private static readonly Color MutedText = Color.FromRgb(120, 120, 130);
        private static readonly Color BorderColor = Color.FromRgb(200, 200, 210);
        private static readonly Color GrayBg = Color.FromRgb(240, 240, 243);
        private static readonly Color WindowBg = Color.FromRgb(245, 245, 248);

        public RefreshZReportWindow(
            string summary, List<string> details)
        {
            Title = "HMV Tools – Refresh Z Report";
            Width = 560;
            Height = 460;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.CanResize;
            Background = new SolidColorBrush(WindowBg);

            var mainGrid = new Grid();
            mainGrid.Margin = new Thickness(20);
            mainGrid.RowDefinitions.Add(
                new RowDefinition { Height = GridLength.Auto });       // 0 Title
            mainGrid.RowDefinitions.Add(
                new RowDefinition { Height = GridLength.Auto });       // 1 Summary
            mainGrid.RowDefinitions.Add(
                new RowDefinition { Height = GridLength.Auto });       // 2 Detail label
            mainGrid.RowDefinitions.Add(
                new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // 3 Detail list
            mainGrid.RowDefinitions.Add(
                new RowDefinition { Height = GridLength.Auto });       // 4 Close btn

            // ── Title ──────────────────────────────────────────────
            var title = new TextBlock
            {
                Text = "Refresh Z — Report",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 12)
            };
            Grid.SetRow(title, 0);
            mainGrid.Children.Add(title);

            // ── Summary ────────────────────────────────────────────
            var summaryBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                Background = new SolidColorBrush(
                    Color.FromRgb(232, 243, 255)),
                Padding = new Thickness(14, 10, 14, 10),
                Margin = new Thickness(0, 0, 0, 12)
            };
            summaryBorder.Child = new TextBlock
            {
                Text = summary,
                FontSize = 13,
                Foreground = new SolidColorBrush(DarkText),
                TextWrapping = TextWrapping.Wrap
            };
            Grid.SetRow(summaryBorder, 1);
            mainGrid.Children.Add(summaryBorder);

            // ── Detail label ───────────────────────────────────────
            var detailLabel = new TextBlock
            {
                Text = "Details",
                FontSize = 12,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(0, 0, 0, 4)
            };
            Grid.SetRow(detailLabel, 2);
            mainGrid.Children.Add(detailLabel);

            // ── Detail list ────────────────────────────────────────
            var listBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 12)
            };
            var listBox = new ListBox
            {
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                FontSize = 12,
                FontFamily = new FontFamily("Consolas"),
                Padding = new Thickness(4)
            };
            foreach (string line in details)
            {
                listBox.Items.Add(new TextBlock
                {
                    Text = line,
                    Padding = new Thickness(6, 4, 6, 4),
                    TextWrapping = TextWrapping.Wrap
                });
            }
            listBorder.Child = listBox;
            Grid.SetRow(listBorder, 3);
            mainGrid.Children.Add(listBorder);

            // ── Close button ───────────────────────────────────────
            var closeBtn = new Button
            {
                Content = "Close",
                Width = 100,
                Height = 36,
                FontSize = 13,
                HorizontalAlignment = HorizontalAlignment.Right,
                Foreground = new SolidColorBrush(
                    Color.FromRgb(60, 60, 60)),
                Background = new SolidColorBrush(GrayBg),
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
            closeBtn.Template = GetRoundButtonTemplate(GrayBg);
            closeBtn.Click += (s, e) => Close();
            Grid.SetRow(closeBtn, 4);
            mainGrid.Children.Add(closeBtn);

            Content = mainGrid;
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
    }
}