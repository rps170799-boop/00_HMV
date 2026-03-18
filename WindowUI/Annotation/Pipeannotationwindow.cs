using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    // ── Enums & Settings ───────────────────────────────────────

    public enum PlacementMode
    {
        GenericAnnotation,
        DetailItem
    }

    public class PipeAnnotationSettings
    {
        public PlacementMode Mode { get; set; }
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
        /// <summary>Only used in GenericAnnotation mode.</summary>
        public double SpacingMm { get; set; }
    }

    // ── Entry container ────────────────────────────────────────

    public class FamilyEntry
    {
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
        public string Display => $"{FamilyName} : {TypeName}";
    }

    // ── Window ─────────────────────────────────────────────────

    public class PipeAnnotationWindow : Window
    {
        // Controls
        private Border annotBtnBorder;
        private Border detailBtnBorder;
        private TextBlock annotBtnText;
        private TextBlock detailBtnText;
        private ListBox listBox;
        private TextBox searchBox;
        private StackPanel spacingPanel;
        private TextBox spacingBox;
        private Border warningBorder;
        private TextBlock subtitleText;
        private TextBlock infoText;

        // Data
        private List<FamilyEntry> annotationEntries;
        private List<FamilyEntry> detailItemEntries;
        private List<FamilyEntry> activeEntries;
        private PlacementMode currentMode = PlacementMode.GenericAnnotation;

        // Colors
        private static readonly Color BluePrimary = Color.FromRgb(0, 120, 212);
        private static readonly Color GrayBg = Color.FromRgb(240, 240, 243);
        private static readonly Color DarkText = Color.FromRgb(30, 30, 30);
        private static readonly Color MutedText = Color.FromRgb(120, 120, 130);
        private static readonly Color BorderColor = Color.FromRgb(200, 200, 210);
        private static readonly Color WarningBg = Color.FromRgb(255, 248, 230);
        private static readonly Color WarningBorder = Color.FromRgb(230, 190, 80);
        private static readonly Color WarningText = Color.FromRgb(140, 100, 20);
        private static readonly Color WindowBg = Color.FromRgb(245, 245, 248);

        /// <summary>User's chosen settings, or null if cancelled.</summary>
        public PipeAnnotationSettings Settings { get; private set; }

        public PipeAnnotationWindow(
            List<FamilyEntry> annotations,
            List<FamilyEntry> detailItems)
        {
            annotationEntries = annotations;
            detailItemEntries = detailItems;
            activeEntries = annotationEntries;

            Title = "HMV Tools - Pipe Annotations";
            Width = 520;
            Height = 580;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(WindowBg);

            var mainGrid = new Grid();
            mainGrid.Margin = new Thickness(20);
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 0 Title
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 1 Mode toggle
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 2 Warning
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 3 Subtitle
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 4 Search
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // 5 List
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 6 Spacing
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 7 Info
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 8 Buttons

            // ── Row 0: Title ───────────────────────────────────────
            var title = new TextBlock
            {
                Text = "Place Elements Along Pipes",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 10)
            };
            Grid.SetRow(title, 0);
            mainGrid.Children.Add(title);

            // ── Row 1: Mode toggle buttons ─────────────────────────
            var modePanel = CreateModeToggle();
            Grid.SetRow(modePanel, 1);
            mainGrid.Children.Add(modePanel);

            // ── Row 2: Warning (Detail Item only) ──────────────────
            warningBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(WarningBorder),
                BorderThickness = new Thickness(1),
                Background = new SolidColorBrush(WarningBg),
                Margin = new Thickness(0, 0, 0, 8),
                Padding = new Thickness(12, 8, 12, 8),
                Visibility = Visibility.Collapsed
            };
            var warningText = new TextBlock
            {
                Text = "⚠  Detail Item mode only works with straight pipes (Lines).\n" +
                       "     Curved pipes (Arcs, Splines) will be skipped.",
                FontSize = 12,
                Foreground = new SolidColorBrush(WarningText),
                TextWrapping = TextWrapping.Wrap
            };
            warningBorder.Child = warningText;
            Grid.SetRow(warningBorder, 2);
            mainGrid.Children.Add(warningBorder);

            // ── Row 3: Subtitle ────────────────────────────────────
            subtitleText = new TextBlock
            {
                Text = "Select a Generic Annotation family type",
                FontSize = 12,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(0, 0, 0, 6)
            };
            Grid.SetRow(subtitleText, 3);
            mainGrid.Children.Add(subtitleText);

            // ── Row 4: Search box ──────────────────────────────────
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
            Grid.SetRow(searchBorder, 4);
            mainGrid.Children.Add(searchBorder);

            // ── Row 5: Family list ─────────────────────────────────
            var listBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 10)
            };
            listBox = new ListBox
            {
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                FontSize = 13,
                Padding = new Thickness(4),
                SelectionMode = SelectionMode.Single
            };
            listBox.MouseDoubleClick += (s, e) =>
            {
                if (listBox.SelectedItem != null) Accept();
            };
            listBorder.Child = listBox;
            Grid.SetRow(listBorder, 5);
            mainGrid.Children.Add(listBorder);

            // ── Row 6: Spacing input (Annotation mode only) ────────
            spacingPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 12)
            };
            spacingPanel.Children.Add(new TextBlock
            {
                Text = "Spacing (mm):",
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0)
            });

            var spacingBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Width = 120
            };
            spacingBox = new TextBox
            {
                Text = "1000",
                Height = 32,
                FontSize = 13,
                VerticalContentAlignment = VerticalAlignment.Center,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(10, 0, 10, 0)
            };
            spacingBox.GotFocus += (s, e) =>
                spacingBorder.BorderBrush = new SolidColorBrush(BluePrimary);
            spacingBox.LostFocus += (s, e) =>
                spacingBorder.BorderBrush = new SolidColorBrush(BorderColor);
            spacingBox.PreviewTextInput += (s, e) =>
            {
                e.Handled = !IsNumericInput(e.Text, spacingBox.Text);
            };
            spacingBorder.Child = spacingBox;
            spacingPanel.Children.Add(spacingBorder);

            Grid.SetRow(spacingPanel, 6);
            mainGrid.Children.Add(spacingPanel);

            // ── Row 7: Info text ───────────────────────────────────
            infoText = new TextBlock
            {
                Text = "Annotations are centered along each pipe segment.\n"
                     + "Orientation follows the pipe tangent direction.\n"
                     + "Works with Pipe, Flex Pipe (straight, arc, spline).",
                FontSize = 11,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(0, 0, 0, 12),
                TextWrapping = TextWrapping.Wrap
            };
            Grid.SetRow(infoText, 7);
            mainGrid.Children.Add(infoText);

            // ── Row 8: Buttons ─────────────────────────────────────
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

            var okBtn = CreateButton("Place Elements", BluePrimary,
                Color.FromRgb(255, 255, 255));
            okBtn.Click += (s, e) => Accept();
            okBtn.Width = 150;

            buttonPanel.Children.Add(cancelBtn);
            buttonPanel.Children.Add(okBtn);
            Grid.SetRow(buttonPanel, 8);
            mainGrid.Children.Add(buttonPanel);

            Content = mainGrid;

            // Initial population
            PopulateList(annotationEntries);

            Loaded += (s, e) => searchBox.Focus();
        }

        // ── Mode toggle ────────────────────────────────────────────

        /// <summary>
        /// Creates the segmented toggle: [Generic Annotation | Detail Item]
        /// </summary>
        private Grid CreateModeToggle()
        {
            var outerBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                Background = new SolidColorBrush(Color.FromRgb(230, 230, 235)),
                Margin = new Thickness(0, 0, 0, 10),
                Padding = new Thickness(3)
            };

            var toggleGrid = new Grid();
            toggleGrid.ColumnDefinitions.Add(
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            toggleGrid.ColumnDefinitions.Add(
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // ─ Left: Generic Annotation (active by default) ────────
            annotBtnBorder = new Border
            {
                CornerRadius = new CornerRadius(6),
                Background = Brushes.White,
                Cursor = Cursors.Hand,
                Margin = new Thickness(0, 0, 2, 0)
            };
            annotBtnText = new TextBlock
            {
                Text = "Generic Annotation",
                FontSize = 13,
                FontWeight = FontWeights.Medium,
                Foreground = new SolidColorBrush(DarkText),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Padding = new Thickness(0, 8, 0, 8)
            };
            annotBtnBorder.Child = annotBtnText;
            annotBtnBorder.MouseLeftButtonDown += (s, e) =>
                SwitchMode(PlacementMode.GenericAnnotation);
            Grid.SetColumn(annotBtnBorder, 0);
            toggleGrid.Children.Add(annotBtnBorder);

            // ─ Right: Detail Item (inactive by default) ────────────
            detailBtnBorder = new Border
            {
                CornerRadius = new CornerRadius(6),
                Background = Brushes.Transparent,
                Cursor = Cursors.Hand,
                Margin = new Thickness(2, 0, 0, 0)
            };
            detailBtnText = new TextBlock
            {
                Text = "Detail Item",
                FontSize = 13,
                FontWeight = FontWeights.Normal,
                Foreground = new SolidColorBrush(MutedText),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Padding = new Thickness(0, 8, 0, 8)
            };
            detailBtnBorder.Child = detailBtnText;
            detailBtnBorder.MouseLeftButtonDown += (s, e) =>
                SwitchMode(PlacementMode.DetailItem);
            Grid.SetColumn(detailBtnBorder, 1);
            toggleGrid.Children.Add(detailBtnBorder);

            outerBorder.Child = toggleGrid;

            // Wrap in a grid row container
            var container = new Grid();
            container.Children.Add(outerBorder);
            return container;
        }

        /// <summary>
        /// Switches the UI between Generic Annotation and Detail Item
        /// modes, updating visuals and the family list.
        /// </summary>
        private void SwitchMode(PlacementMode mode)
        {
            if (mode == currentMode) return;
            currentMode = mode;

            // Clear search
            searchBox.Text = "";

            if (mode == PlacementMode.GenericAnnotation)
            {
                // Toggle visuals
                annotBtnBorder.Background = Brushes.White;
                annotBtnText.FontWeight = FontWeights.Medium;
                annotBtnText.Foreground = new SolidColorBrush(DarkText);

                detailBtnBorder.Background = Brushes.Transparent;
                detailBtnText.FontWeight = FontWeights.Normal;
                detailBtnText.Foreground = new SolidColorBrush(MutedText);

                // Show/hide controls
                warningBorder.Visibility = Visibility.Collapsed;
                spacingPanel.Visibility = Visibility.Visible;

                // Update texts
                subtitleText.Text = "Select a Generic Annotation family type";
                infoText.Text =
                    "Annotations are centered along each pipe segment.\n"
                    + "Orientation follows the pipe tangent direction.\n"
                    + "Works with Pipe, Flex Pipe (straight, arc, spline).";

                activeEntries = annotationEntries;
                PopulateList(annotationEntries);
            }
            else
            {
                // Toggle visuals
                detailBtnBorder.Background = Brushes.White;
                detailBtnText.FontWeight = FontWeights.Medium;
                detailBtnText.Foreground = new SolidColorBrush(DarkText);

                annotBtnBorder.Background = Brushes.Transparent;
                annotBtnText.FontWeight = FontWeights.Normal;
                annotBtnText.Foreground = new SolidColorBrush(MutedText);

                // Show/hide controls
                warningBorder.Visibility = Visibility.Visible;
                spacingPanel.Visibility = Visibility.Collapsed;

                // Update texts
                subtitleText.Text = "Select a Detail Item family type";
                infoText.Text =
                    "A single instance is placed spanning the full pipe length.\n"
                    + "The detail item stretches from start to end point.\n"
                    + "Only works with line-based Detail Item families.";

                activeEntries = detailItemEntries;
                PopulateList(detailItemEntries);
            }
        }

        // ── Accept logic ───────────────────────────────────────────

        private void Accept()
        {
            if (listBox.SelectedItem == null)
            {
                string what = currentMode == PlacementMode.GenericAnnotation
                    ? "annotation" : "detail item";
                MessageBox.Show($"Select a {what} family.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            // Validate spacing only for annotation mode
            double spacing = 0;
            if (currentMode == PlacementMode.GenericAnnotation)
            {
                if (!double.TryParse(spacingBox.Text, out spacing)
                    || spacing <= 0)
                {
                    MessageBox.Show(
                        "Enter a valid spacing value greater than 0.",
                        "HMV Tools", MessageBoxButton.OK);
                    return;
                }
            }

            // Resolve selected entry
            string selectedText =
                (listBox.SelectedItem as TextBlock)?.Text ?? "";
            var entry = activeEntries.FirstOrDefault(
                e => e.Display == selectedText);

            if (entry == null)
            {
                MessageBox.Show("Could not resolve selected family.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            Settings = new PipeAnnotationSettings
            {
                Mode = currentMode,
                FamilyName = entry.FamilyName,
                TypeName = entry.TypeName,
                SpacingMm = spacing
            };

            DialogResult = true;
            Close();
        }

        // ── List population ────────────────────────────────────────

        private void PopulateList(List<FamilyEntry> entries)
        {
            listBox.Items.Clear();
            foreach (var entry in entries)
                listBox.Items.Add(CreateListItem(entry.Display));

            if (listBox.Items.Count > 0)
                listBox.SelectedIndex = 0;
        }

        // ── Input validation ───────────────────────────────────────

        private bool IsNumericInput(string newText, string currentText)
        {
            foreach (char c in newText)
            {
                if (!char.IsDigit(c) && c != '.')
                    return false;
                if (c == '.' && currentText.Contains("."))
                    return false;
            }
            return true;
        }

        // ── Search filter ──────────────────────────────────────────

        private void SearchBox_TextChanged(object sender,
            TextChangedEventArgs e)
        {
            listBox.Items.Clear();
            string filter = searchBox.Text.ToLower();
            foreach (var entry in activeEntries)
            {
                if (entry.Display.ToLower().Contains(filter))
                    listBox.Items.Add(CreateListItem(entry.Display));
            }
        }

        // ── UI helpers (same pattern as DwgConvertWindow) ──────────

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
}