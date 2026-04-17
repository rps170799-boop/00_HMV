using Autodesk.Revit.DB;
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
        private string _status = "Not Linked";

        public static readonly string[] PlacementOptions = new[]
        {
            "Shared Coordinates",
            "Origin to Origin",
            "Center to Center"
        };

        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); }
        }

        public string Name { get; set; }
        public string FolderPath { get; set; }
        public string Urn { get; set; }
        public int LinkTypeElementId { get; set; }


        public string Status
        {
            get => _status;
            set { _status = value; OnPropertyChanged(nameof(Status)); }
        }

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
                    IsSelected = false,
                    Status = f.Status,
                    LinkTypeElementId = f.LinkTypeElementId
                };
                row.PropertyChanged += Row_PropertyChanged;
                FileRows.Add(row);
            }

            // Apply row colors after grid is loaded
            FileGrid.LoadingRow += FileGrid_LoadingRow;

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
            SelectedFiles = FileRows.Where(r => r.IsSelected && r.Status == "Not Linked").ToList();

            if (SelectedFiles.Count == 0)
            {
                MessageBox.Show("No unlinked files selected.\nOnly 'Not Linked' files can be linked.\nUse 'Reload Selected' for existing links.",
                    "HMV Tools");
                return;
            }
            DialogResult = true;
            Close();
        }

        private void FileGrid_LoadingRow(object sender, System.Windows.Controls.DataGridRowEventArgs e)
        {
            if (e.Row.Item is RvtFileRow row)
            {
                switch (row.Status)
                {
                    case "Loaded":
                        e.Row.Background = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(220, 237, 220)); // light green
                        break;
                    case "Unloaded":
                        e.Row.Background = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(255, 243, 205)); // light yellow
                        break;
                    default:
                        e.Row.Background = System.Windows.Media.Brushes.Transparent;
                        break;
                }
            }
        }

        /// <summary>
        /// Files selected for reloading. Populated when user clicks "Reload Selected".
        /// </summary>
        public List<RvtFileRow> ReloadFiles { get; private set; }
        public bool IsReloadAction { get; private set; } = false;

        private void ReloadSelected_Click(object sender, RoutedEventArgs e)
        {
            ReloadFiles = FileRows.Where(r => r.IsSelected &&
                (r.Status == "Loaded" || r.Status == "Unloaded")).ToList();

            if (ReloadFiles.Count == 0)
            {
                MessageBox.Show("No linked files selected.\nSelect files with status 'Loaded' or 'Unloaded'.",
                    "HMV Tools");
                return;
            }

            IsReloadAction = true;
            DialogResult = true;
            Close();
        }
        private void SearchBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            string filter = SearchBox.Text?.Trim() ?? "";
            if (string.IsNullOrEmpty(filter))
            {
                FileGrid.ItemsSource = FileRows;
            }
            else
            {
                var filtered = FileRows.Where(r =>
                    r.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    r.FolderPath.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0
                ).ToList();
                FileGrid.ItemsSource = filtered;
            }
        }
        private void TopBar_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
                DragMove();
        }
        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}