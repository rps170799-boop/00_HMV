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
        public int View3DIndex { get; set; }
        public bool HasShoulder { get; set; }
    }

    // ── Window Code-Behind ─────────────────────────────────────
    public partial class SpotElevationWindow : Window
    {
        private List<LinkInfo> linkInfos;
        private List<string> foundationSourceItems;
        private List<string> floorLinkItems;
        private List<string> view3DItems;

        public SpotElevationSettings Settings { get; private set; }

        public SpotElevationWindow(List<LinkInfo> links, List<string> viewNames)
        {
            InitializeComponent();

            linkInfos = links;
            view3DItems = viewNames;
            floorLinkItems = links.Select(l => l.Name).ToList();
            foundationSourceItems = BuildFoundationSourceItems();

            // Populate Comboboxes initially
            PopulateCombo(cmbFoundationSource, foundationSourceItems);
            if (foundationSourceItems.Count > 0) cmbFoundationSource.SelectedIndex = 0;

            PopulateCombo(cmbFloorLink, floorLinkItems);
            if (floorLinkItems.Count > 0) cmbFloorLink.SelectedIndex = 0;

            PopulateCombo(cmbView3D, view3DItems);
            if (view3DItems.Count > 0) cmbView3D.SelectedIndex = 0;
        }

        // ── Custom Title Bar Logic ────────────────────────────────

        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        // ── Initialization & Filtering ────────────────────────────

        private List<string> BuildFoundationSourceItems()
        {
            var items = new List<string> { "Active Model" };
            items.AddRange(linkInfos.Select(l => l.Name));
            return items;
        }

        private void PopulateCombo(ComboBox combo, List<string> items)
        {
            combo.Items.Clear();
            foreach (var item in items)
            {
                combo.Items.Add(item);
            }
        }

        private void FilterCombo(ComboBox combo, List<string> allItems, string filterText)
        {


            if (combo == null || allItems == null) return;
            string selected = combo.SelectedItem as string;
            string filter = (filterText == "Search...") ? "" : filterText;

            combo.Items.Clear();
            var filtered = string.IsNullOrWhiteSpace(filter)
                ? allItems
                : allItems.Where(i => i.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0).ToList();

            foreach (var item in filtered)
            {
                combo.Items.Add(item);
            }

            if (selected != null && filtered.Contains(selected))
            {
                combo.SelectedItem = selected;
            }
            else if (filtered.Count == 1)
            {
                combo.SelectedIndex = 0;
            }
        }

        private void SearchFoundation_TextChanged(object sender, TextChangedEventArgs e)
        {
            FilterCombo(cmbFoundationSource, foundationSourceItems, txtSearchFoundation.Text);
        }

        private void SearchFloor_TextChanged(object sender, TextChangedEventArgs e)
        {
            FilterCombo(cmbFloorLink, floorLinkItems, txtSearchFloor.Text);
        }

        private void SearchView_TextChanged(object sender, TextChangedEventArgs e)
        {
            FilterCombo(cmbView3D, view3DItems, txtSearchView.Text);
        }

        // ── UI Visual Handlers ────────────────────────────────────

        private void SearchBox_GotFocus(object sender, RoutedEventArgs e)
        {
            var searchBox = sender as TextBox;
            if (searchBox != null && searchBox.Text == "Search...")
            {
                searchBox.Text = "";
                searchBox.Foreground = new SolidColorBrush(Color.FromRgb(30, 30, 30)); // DarkText
            }

            if (searchBox == txtSearchFoundation) brdSearchFoundation.BorderBrush = new SolidColorBrush(Color.FromRgb(0, 73, 130));
            if (searchBox == txtSearchFloor) brdSearchFloor.BorderBrush = new SolidColorBrush(Color.FromRgb(0, 73, 130));
            if (searchBox == txtSearchView) brdSearchView.BorderBrush = new SolidColorBrush(Color.FromRgb(0, 73, 130));
        }

        private void SearchBox_LostFocus(object sender, RoutedEventArgs e)
        {
            var searchBox = sender as TextBox;
            if (searchBox != null && string.IsNullOrWhiteSpace(searchBox.Text))
            {
                searchBox.Text = "Search...";
                searchBox.Foreground = new SolidColorBrush(Color.FromRgb(120, 120, 130)); // MutedText
            }

            if (searchBox == txtSearchFoundation) brdSearchFoundation.BorderBrush = new SolidColorBrush(Color.FromRgb(200, 200, 210));
            if (searchBox == txtSearchFloor) brdSearchFloor.BorderBrush = new SolidColorBrush(Color.FromRgb(200, 200, 210));
            if (searchBox == txtSearchView) brdSearchView.BorderBrush = new SolidColorBrush(Color.FromRgb(200, 200, 210));
        }

        private void TxtLeader_GotFocus(object sender, RoutedEventArgs e)
        {
            brdLeaderOffset.BorderBrush = new SolidColorBrush(Color.FromRgb(0, 73, 130));
        }

        private void TxtLeader_LostFocus(object sender, RoutedEventArgs e)
        {
            brdLeaderOffset.BorderBrush = new SolidColorBrush(Color.FromRgb(200, 200, 210));
        }

        private void NumericOnly(object sender, TextCompositionEventArgs e)
        {
            foreach (char c in e.Text)
            {
                if (!char.IsDigit(c) && c != '.') { e.Handled = true; return; }
                if (c == '.' && txtLeaderOffset.Text.Contains(".")) { e.Handled = true; return; }
            }
        }

        // ── Action Buttons ────────────────────────────────────────

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            if (cmbFloorLink.SelectedIndex < 0)
            {
                MessageBox.Show("Select the link that contains the floors.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!double.TryParse(txtLeaderOffset.Text, out double offset) || offset <= 0)
            {
                MessageBox.Show("Enter a valid leader offset > 0.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Resolve indices
            string foundText = cmbFoundationSource.SelectedItem as string;
            int foundSrcIdx = foundText != null ? foundationSourceItems.IndexOf(foundText) - 1 : -1;

            string floorText = cmbFloorLink.SelectedItem as string;
            int floorLinkIdx = floorText != null ? floorLinkItems.IndexOf(floorText) : -1;

            if (floorLinkIdx < 0)
            {
                MessageBox.Show("Select a valid floor link.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string viewText = cmbView3D.SelectedItem as string;
            int viewIdx = viewText != null ? view3DItems.IndexOf(viewText) : -1;

            if (viewIdx < 0)
            {
                MessageBox.Show("Select a valid 3D View.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
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
                View3DIndex = viewIdx,
                HasShoulder = chkShoulder.IsChecked == true
            };

            DialogResult = true;
            Close();
        }
    }
}