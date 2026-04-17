using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;

namespace HMVTools
{
    /// <summary>
    /// View model for each .rvt file row in the DataGrid.
    /// </summary>
    public class RvtFileRow : INotifyPropertyChanged
    {
        private bool _isSelected;
        private string _placement = "Shared Coordinates";

        public static readonly string[] PlacementOptions = new[]
        {
            "Shared Coordinates",
            "Origin to Origin",
            "Center to Center",
            "By Shared Coordinates"
        };

        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); }
        }

        public string Name { get; set; }
        public string FolderPath { get; set; }
        public string Urn { get; set; }

        public string Placement
        {
            get => _placement;
            set { _placement = value; OnPropertyChanged(nameof(Placement)); }
        }

        public string[] GetPlacementOptions() => PlacementOptions;

        // Expose as instance property for XAML binding
        public IEnumerable<string> PlacementOptionsBinding => PlacementOptions;

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public partial class BatchCloudLinkWindow : Window
    {
        public ObservableCollection<RvtFileRow> FileRows { get; private set; }

        /// <summary>
        /// The files the user selected for linking. Populated when they click "Link Selected".
        /// </summary>
        public List<RvtFileRow> SelectedFiles { get; private set; }

        public BatchCloudLinkWindow(
            List<CloudRevitFile> rvtFiles,
            string projectName,
            string folderName)
        {
            InitializeComponent();

            // Build rows from scan results
            FileRows = new ObservableCollection<RvtFileRow>();
            foreach (var f in rvtFiles)
            {
                var row = new RvtFileRow
                {
                    Name = f.Name,
                    FolderPath = string.IsNullOrEmpty(f.Path) ? "(root)" : f.Path,
                    Urn = f.Urn,
                    IsSelected = false
                };
                row.PropertyChanged += Row_PropertyChanged;
                FileRows.Add(row);
            }

            FileGrid.ItemsSource = FileRows;

            // Header
            HeaderLabel.Text = $"Found {rvtFiles.Count} RVT file(s)";
            SubHeaderLabel.Text = $"{projectName} / {folderName}";

            // Bulk placement combo
            foreach (string opt in RvtFileRow.PlacementOptions)
                BulkPlacementCombo.Items.Add(opt);
            BulkPlacementCombo.SelectedIndex = 0;

            UpdateCount();
        }

        // ─── Workaround: PlacementOptions binding for DataGrid ComboBox ───
        // DataGridTemplateColumn can't resolve instance properties easily,
        // so we use a resource-based approach via the Loaded event.
        // The XAML binding "PlacementOptions" needs a source.
        // We'll fix this by overriding the ItemsSource in code.

        private void Row_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(RvtFileRow.IsSelected))
                UpdateCount();
        }

        private void UpdateCount()
        {
            int count = FileRows.Count(r => r.IsSelected);
            CountLabel.Text = $"{count} selected";
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var row in FileRows) row.IsSelected = true;
        }

        private void ClearSelection_Click(object sender, RoutedEventArgs e)
        {
            foreach (var row in FileRows) row.IsSelected = false;
        }

        private void ApplyPlacement_Click(object sender, RoutedEventArgs e)
        {
            string selected = BulkPlacementCombo.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(selected)) return;

            foreach (var row in FileRows.Where(r => r.IsSelected))
                row.Placement = selected;
        }

        private void LinkSelected_Click(object sender, RoutedEventArgs e)
        {
            SelectedFiles = FileRows.Where(r => r.IsSelected).ToList();

            if (SelectedFiles.Count == 0)
            {
                MessageBox.Show("No files selected.", "HMV Tools");
                return;
            }

            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}