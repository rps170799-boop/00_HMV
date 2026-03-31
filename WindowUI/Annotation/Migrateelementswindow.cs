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

    /// <summary>Represents an open Revit document for the target dropdown.</summary>
    public class OpenDocEntry
    {
        public string Title { get; set; }
        public string PathName { get; set; }
        /// <summary>Index into Application.Documents for retrieval.</summary>
        public int Index { get; set; }
    }

    /// <summary>Represents a view available for transfer.</summary>
    public class ViewEntry
    {
        public int Id { get; set; }
        public string Name { get; set; }
        /// <summary>Grouping label: "Floor Plans", "Sections", "Drafting Views", "Legends".</summary>
        public string Category { get; set; }
    }

    /// <summary>User selections returned to the command.</summary>
    public class MigrationSettings
    {
        public int TargetDocIndex { get; set; }
        public List<int> SelectedViewIds { get; set; } = new List<int>();
        public bool IncludeAnnotations { get; set; } = true;
        public bool IncludeRefMarkers { get; set; } = true;
    }

    // ── Window ─────────────────────────────────────────────────

    public class MigrateElementsWindow : Window
    {
        // ── Colors (same palette as PipeAnnotationWindow) ──────
        private static readonly Color BluePrimary = Color.FromRgb(0, 120, 212);
        private static readonly Color GrayBg = Color.FromRgb(240, 240, 243);
        private static readonly Color DarkText = Color.FromRgb(30, 30, 30);
        private static readonly Color MutedText = Color.FromRgb(120, 120, 130);
        private static readonly Color BorderColor = Color.FromRgb(200, 200, 210);
        private static readonly Color WindowBg = Color.FromRgb(245, 245, 248);
        private static readonly Color AccentLight = Color.FromRgb(232, 243, 255);
        private static readonly Color SectionHead = Color.FromRgb(60, 60, 70);

        // ── Controls ───────────────────────────────────────────
        private ComboBox targetDocCombo;
        private StackPanel viewListPanel;
        private ScrollViewer viewScrollViewer;
        private CheckBox chkAnnotations;
        private CheckBox chkRefMarkers;
        private TextBlock statusText;

        // ── Data ───────────────────────────────────────────────
        private readonly List<OpenDocEntry> openDocs;
        private readonly List<ViewEntry> sourceViews;
        private readonly string sourceTitle;
        private readonly int elementCount;

        /// <summary>Maps ViewEntry.Id → CheckBox for retrieval.</summary>
        private readonly Dictionary<int, CheckBox> viewCheckBoxes
            = new Dictionary<int, CheckBox>();

        /// <summary>Result — null if cancelled.</summary>
        public MigrationSettings Settings { get; private set; }

        // ── Constructor ────────────────────────────────────────

        public MigrateElementsWindow(
            string sourceDocTitle,
            int selectedElementCount,
            List<OpenDocEntry> targetDocs,
            List<ViewEntry> views)
        {
            sourceTitle = sourceDocTitle;
            elementCount = selectedElementCount;
            openDocs = targetDocs;
            sourceViews = views;

            Title = "HMV Tools – Migrate Elements";
            Width = 560;
            Height = 660;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(WindowBg);

            Content = BuildLayout();
        }

        // ── Layout ─────────────────────────────────────────────

        private Grid BuildLayout()
        {
            var root = new Grid { Margin = new Thickness(22) };

            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                     // 0 Title
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                     // 1 Source info
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                     // 2 Target combo
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                     // 3 "Views" label
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });// 4 View list
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                     // 5 Options
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                     // 6 Status
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                     // 7 Buttons

            // ── 0: Title ───────────────────────────────────────
            var title = new TextBlock
            {
                Text = "Migrate Elements to Target",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 6)
            };
            Grid.SetRow(title, 0);
            root.Children.Add(title);

            // ── 1: Source info bar ─────────────────────────────
            var infoBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                Background = new SolidColorBrush(AccentLight),
                Padding = new Thickness(12, 8, 12, 8),
                Margin = new Thickness(0, 0, 0, 12)
            };
            infoBorder.Child = new TextBlock
            {
                Text = $"Source:  {sourceTitle}   ·   {elementCount} element{(elementCount != 1 ? "s" : "")} selected",
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(30, 80, 160))
            };
            Grid.SetRow(infoBorder, 1);
            root.Children.Add(infoBorder);

            // ── 2: Target document combo ───────────────────────
            var targetLabel = new TextBlock
            {
                Text = "Target Document",
                FontSize = 12,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(0, 0, 0, 4)
            };

            var comboBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 14)
            };

            targetDocCombo = new ComboBox
            {
                Height = 34,
                FontSize = 13,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(8, 0, 8, 0),
                VerticalContentAlignment = VerticalAlignment.Center
            };
            foreach (var doc in openDocs)
                targetDocCombo.Items.Add(doc.Title);
            if (targetDocCombo.Items.Count > 0)
                targetDocCombo.SelectedIndex = 0;

            comboBorder.Child = targetDocCombo;

            var targetStack = new StackPanel { Margin = new Thickness(0) };
            targetStack.Children.Add(targetLabel);
            targetStack.Children.Add(comboBorder);
            Grid.SetRow(targetStack, 2);
            root.Children.Add(targetStack);

            // ── 3: "Views to Transfer" label ───────────────────
            var viewsLabel = new TextBlock
            {
                Text = "Views to Transfer",
                FontSize = 12,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(0, 0, 0, 4)
            };
            Grid.SetRow(viewsLabel, 3);
            root.Children.Add(viewsLabel);

            // ── 4: View checklist (scrollable) ─────────────────
            var listBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 12)
            };

            viewListPanel = new StackPanel { Margin = new Thickness(4) };
            viewScrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = viewListPanel
            };
            listBorder.Child = viewScrollViewer;
            Grid.SetRow(listBorder, 4);
            root.Children.Add(listBorder);

            PopulateViewList();

            // ── 5: Options ─────────────────────────────────────
            var optPanel = new StackPanel { Margin = new Thickness(0, 0, 0, 8) };

            chkAnnotations = new CheckBox
            {
                Content = "  Include view-specific annotations (2-step copy)",
                IsChecked = true,
                FontSize = 12,
                Margin = new Thickness(0, 0, 0, 4)
            };
            chkRefMarkers = new CheckBox
            {
                Content = "  Recreate reference markers (sections / callouts)",
                IsChecked = true,
                FontSize = 12
            };
            optPanel.Children.Add(chkAnnotations);
            optPanel.Children.Add(chkRefMarkers);

            Grid.SetRow(optPanel, 5);
            root.Children.Add(optPanel);

            // ── 6: Status line ─────────────────────────────────
            statusText = new TextBlock
            {
                Text = "",
                FontSize = 11,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(0, 0, 0, 10)
            };
            Grid.SetRow(statusText, 6);
            root.Children.Add(statusText);

            UpdateStatus();

            // ── 7: Buttons ─────────────────────────────────────
            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var cancelBtn = CreateButton("Cancel", GrayBg, Color.FromRgb(60, 60, 60));
            cancelBtn.Width = 90;
            cancelBtn.Margin = new Thickness(0, 0, 8, 0);
            cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };

            var migrateBtn = CreateButton("Migrate", BluePrimary, Colors.White);
            migrateBtn.Width = 140;
            migrateBtn.Click += (s, e) => Accept();

            btnPanel.Children.Add(cancelBtn);
            btnPanel.Children.Add(migrateBtn);
            Grid.SetRow(btnPanel, 7);
            root.Children.Add(btnPanel);

            return root;
        }

        // ── View list population ───────────────────────────────

        private void PopulateViewList()
        {
            viewListPanel.Children.Clear();
            viewCheckBoxes.Clear();

            // Group by category, preserve order
            var groups = sourceViews
                .GroupBy(v => v.Category)
                .OrderBy(g => CategorySortOrder(g.Key));

            foreach (var group in groups)
            {
                // ── Category header with "select all" toggle ───
                var headerCheck = new CheckBox
                {
                    IsChecked = false,
                    FontSize = 13,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush(SectionHead),
                    Margin = new Thickness(4, 8, 0, 2),
                    Content = $"  {group.Key}  ({group.Count()})"
                };

                // Capture for closure
                var categoryViews = group.ToList();
                headerCheck.Checked += (s, e) => SetCategoryChecked(categoryViews, true);
                headerCheck.Unchecked += (s, e) => SetCategoryChecked(categoryViews, false);

                viewListPanel.Children.Add(headerCheck);

                // ── Individual view checkboxes ─────────────────
                foreach (var view in group.OrderBy(v => v.Name))
                {
                    var cb = new CheckBox
                    {
                        Content = $"  {view.Name}",
                        Tag = view.Id,
                        FontSize = 12,
                        Margin = new Thickness(22, 2, 0, 2),
                        Foreground = new SolidColorBrush(DarkText)
                    };
                    cb.Checked += (s, e) => UpdateStatus();
                    cb.Unchecked += (s, e) => UpdateStatus();

                    viewCheckBoxes[view.Id] = cb;
                    viewListPanel.Children.Add(cb);
                }
            }
        }

        private void SetCategoryChecked(List<ViewEntry> views, bool isChecked)
        {
            foreach (var v in views)
            {
                if (viewCheckBoxes.TryGetValue(v.Id, out CheckBox cb))
                    cb.IsChecked = isChecked;
            }
            UpdateStatus();
        }

        private void UpdateStatus()
        {
            int count = viewCheckBoxes.Values.Count(cb => cb.IsChecked == true);
            statusText.Text = count == 0
                ? "No views selected — only 3D elements will be copied."
                : $"{count} view{(count != 1 ? "s" : "")} selected for transfer.";
        }

        private int CategorySortOrder(string cat)
        {
            switch (cat)
            {
                case "Floor Plans": return 0;
                case "Ceiling Plans": return 1;
                case "Sections": return 2;
                case "Drafting Views": return 3;
                case "Legends": return 4;
                default: return 9;
            }
        }

        // ── Accept ─────────────────────────────────────────────

        private void Accept()
        {
            if (targetDocCombo.SelectedIndex < 0)
            {
                MessageBox.Show("Select a target document.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            var selectedIds = viewCheckBoxes
                .Where(kv => kv.Value.IsChecked == true)
                .Select(kv => kv.Key)
                .ToList();

            Settings = new MigrationSettings
            {
                TargetDocIndex = targetDocCombo.SelectedIndex,
                SelectedViewIds = selectedIds,
                IncludeAnnotations = chkAnnotations.IsChecked == true,
                IncludeRefMarkers = chkRefMarkers.IsChecked == true
            };

            DialogResult = true;
            Close();
        }

        // ── UI helpers (same as PipeAnnotationWindow) ──────────

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