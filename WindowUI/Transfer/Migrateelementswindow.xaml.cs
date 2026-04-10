using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    // Data classes (OpenDocEntry, ViewEntry, MigrationSettings)
    // are defined in MigrationDataClasses.cs

    public partial class MigrateElementsWindow : Window
    {
        private static readonly Color DarkText = Color.FromRgb(30, 30, 30);
        private static readonly Color SectionHead = Color.FromRgb(60, 60, 70);

        private readonly List<OpenDocEntry> _openDocs;
        private readonly List<ViewEntry> _sourceViews;
        private readonly string _sourceTitle;
        private readonly int _elementCount;

        private readonly Dictionary<int, CheckBox> _viewCheckBoxes
            = new Dictionary<int, CheckBox>();

        public MigrationSettings Settings { get; private set; }

        public MigrateElementsWindow(
            string sourceDocTitle,
            int selectedElementCount,
            List<OpenDocEntry> targetDocs,
            List<ViewEntry> views)
        {
            InitializeComponent();

            _sourceTitle = sourceDocTitle;
            _elementCount = selectedElementCount;
            _openDocs = targetDocs ?? new List<OpenDocEntry>();
            _sourceViews = views ?? new List<ViewEntry>();

            // Source info line
            txtSourceInfo.Text = _elementCount > 0
                ? $"Source:  {_sourceTitle}   ·   {_elementCount} element{(_elementCount != 1 ? "s" : "")} selected"
                : $"Source:  {_sourceTitle}   ·   Views-only migration";

            // Populate target docs
            foreach (var doc in _openDocs) targetDocCombo.Items.Add(doc.Title);
            if (targetDocCombo.Items.Count > 0) targetDocCombo.SelectedIndex = 0;

            // Populate view list
            PopulateViewList();
            UpdateStatus();
        }

        // ── View list population (preserved from original) ─────────
        private void PopulateViewList()
        {
            viewListPanel.Children.Clear();
            _viewCheckBoxes.Clear();

            var groups = _sourceViews
                .GroupBy(v => v.Category)
                .OrderBy(g => CategoryOrder(g.Key));

            foreach (var grp in groups)
            {
                var catViews = grp.ToList();

                var header = new CheckBox
                {
                    IsChecked = false,
                    FontSize = 12,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush(SectionHead),
                    Margin = new Thickness(2, 6, 0, 2),
                    Content = $"  {grp.Key}  ({catViews.Count})"
                };
                header.Checked += (s, e) => SetCat(catViews, true);
                header.Unchecked += (s, e) => SetCat(catViews, false);
                viewListPanel.Children.Add(header);

                foreach (var v in catViews.OrderBy(x => x.Name))
                {
                    var cb = new CheckBox
                    {
                        Content = $"  {v.Name}",
                        Tag = v.Id,
                        FontSize = 11,
                        Margin = new Thickness(20, 2, 0, 2),
                        Foreground = new SolidColorBrush(DarkText)
                    };
                    cb.Checked += (s, e) => UpdateStatus();
                    cb.Unchecked += (s, e) => UpdateStatus();
                    _viewCheckBoxes[v.Id] = cb;
                    viewListPanel.Children.Add(cb);
                }
            }
        }

        // ── Search filter ──────────────────────────────────────────
        private void ViewSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyViewFilter();
        }

        private void ApplyViewFilter()
        {
            string q = (viewSearchBox.Text ?? string.Empty).Trim();
            bool empty = q.Length == 0;

            CheckBox currentHeader = null;
            int visibleInGroup = 0;

            var children = viewListPanel.Children;
            for (int i = 0; i < children.Count; i++)
            {
                var cb = children[i] as CheckBox;
                if (cb == null) continue;

                if (cb.Tag == null)
                {
                    // Category header — finalize previous group
                    if (currentHeader != null)
                        currentHeader.Visibility = visibleInGroup > 0
                            ? System.Windows.Visibility.Visible
                            : System.Windows.Visibility.Collapsed;

                    currentHeader = cb;
                    visibleInGroup = 0;
                }
                else
                {
                    string name = cb.Content?.ToString() ?? string.Empty;
                    bool match = empty
                        || name.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0;
                    cb.Visibility = match
                        ? System.Windows.Visibility.Visible
                        : System.Windows.Visibility.Collapsed;
                    if (match) visibleInGroup++;
                }
            }

            if (currentHeader != null)
                currentHeader.Visibility = visibleInGroup > 0
                    ? System.Windows.Visibility.Visible
                    : System.Windows.Visibility.Collapsed;
        }

        private void SetCat(List<ViewEntry> views, bool val)
        {
            foreach (var v in views)
                if (_viewCheckBoxes.TryGetValue(v.Id, out var cb))
                    cb.IsChecked = val;
            UpdateStatus();
        }

        private void UpdateStatus()
        {
            if (statusText == null) return;  // safety during construction
            int n = _viewCheckBoxes.Values.Count(c => c.IsChecked == true);
            statusText.Text = n == 0
                ? "No views selected — only 3D elements will be copied."
                : $"{n} view{(n != 1 ? "s" : "")} selected for transfer.";
        }

        private int CategoryOrder(string c)
        {
            switch (c)
            {
                case "Floor Plans": return 0;
                case "Ceiling Plans": return 1;
                case "Sections": return 2;
                case "Drafting Views": return 3;
                case "Legends": return 4;
                default: return 9;
            }
        }

        // ── Title bar drag / close ─────────────────────────────────
        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        // ── Migrate / validate ─────────────────────────────────────
        private void BtnMigrate_Click(object sender, RoutedEventArgs e)
        {
            if (targetDocCombo.SelectedIndex < 0)
            {
                MessageBox.Show("Select a target document.", "HMV Tools",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Settings = new MigrationSettings
            {
                TargetDocIndex = targetDocCombo.SelectedIndex,
                SelectedViewIds = _viewCheckBoxes
                    .Where(kv => kv.Value.IsChecked == true)
                    .Select(kv => kv.Key)
                    .ToList(),
                IncludeAnnotations = chkAnnotations.IsChecked == true,
                IncludeRefMarkers = chkRefMarkers.IsChecked == true
            };

            DialogResult = true;
            Close();
        }
    }
}