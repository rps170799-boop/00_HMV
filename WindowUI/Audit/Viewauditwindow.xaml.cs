using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace HMVTools
{
    public class ViewAuditEntry : INotifyPropertyChanged
    {
        public int ElementId { get; set; }
        public string ViewType { get; set; }
        public string OriginalName { get; set; }
        public bool HasConflict { get; set; }

        private string _newName;
        public string NewName
        {
            get => _newName;
            set
            {
                if (_newName != value)
                {
                    _newName = value;
                    OnPropertyChanged(nameof(NewName));
                    OnPropertyChanged(nameof(NameChanged));
                }
            }
        }

        public string Sheets { get; set; }
        public int SheetCount { get; set; }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        public bool NameChanged => OriginalName != NewName;

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
        }
    }

    public partial class ViewAuditWindow : Window
    {
        private ObservableCollection<ViewAuditEntry> AllEntries;
        private ObservableCollection<ViewAuditEntry> FilteredEntries;
        private HashSet<string> AllViewTypeNames;
        private int lastCheckedIndex = -1; // Memory for Shift+Click selection

        public List<ViewAuditEntry> Results => AllEntries.ToList();

        public ViewAuditWindow(List<ViewAuditEntry> data, HashSet<string> allViewTypeNames)
        {
            InitializeComponent();

            AllViewTypeNames = allViewTypeNames;
            AllEntries = new ObservableCollection<ViewAuditEntry>(data);
            FilteredEntries = new ObservableCollection<ViewAuditEntry>(AllEntries);

            dgViews.ItemsSource = FilteredEntries;

            foreach (var item in AllEntries)
            {
                item.PropertyChanged += Entry_PropertyChanged;
            }

            UpdateStatus();
            this.Loaded += (s, e) => txtSearch.Focus();
        }

        // ── Top Bar Logic ──
        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            dgViews.CommitEdit();
            dgViews.CommitEdit();

            var modifiedCount = AllEntries.Count(x => x.NameChanged);
            if (modifiedCount == 0)
            {
                MessageBox.Show("No view names have been modified.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            this.DialogResult = true;
            this.Close();
        }

        // ── Independent Mass Edit Logic ──
        private void BtnApplyPrefix_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = FilteredEntries.Where(x => x.IsSelected).ToList();
            if (selectedItems.Count == 0)
            {
                MessageBox.Show("Please select at least one view.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string prefixInput = txtPrefix.Text;
            if (string.IsNullOrEmpty(prefixInput))
            {
                MessageBox.Show("Please enter a valid prefix.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            foreach (var item in selectedItems)
            {
                item.NewName = prefixInput + item.NewName; // Applies to current NewName to allow stacking operations
            }
        }

        private void BtnApplyCut_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = FilteredEntries.Where(x => x.IsSelected).ToList();
            if (selectedItems.Count == 0)
            {
                MessageBox.Show("Please select at least one view.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string cutInput = txtCut.Text;
            if (string.IsNullOrEmpty(cutInput))
            {
                MessageBox.Show("Please enter the exact text you want to remove.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            foreach (var item in selectedItems)
            {
                if (item.NewName.Contains(cutInput))
                {
                    item.NewName = item.NewName.Replace(cutInput, "");
                }
            }
        }

        // ── Shift+Click Multiselect Logic ──
        private void CheckBox_Click(object sender, RoutedEventArgs e)
        {
            var cb = sender as CheckBox;
            var row = DataGridRow.GetRowContainingElement(cb);
            if (row == null) return;

            int currentIndex = row.GetIndex();
            bool isShiftPressed = Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift);
            bool targetState = cb.IsChecked == true;

            if (isShiftPressed && lastCheckedIndex != -1)
            {
                int start = Math.Min(lastCheckedIndex, currentIndex);
                int end = Math.Max(lastCheckedIndex, currentIndex);

                for (int i = start; i <= end; i++)
                {
                    if (i < FilteredEntries.Count)
                    {
                        FilteredEntries[i].IsSelected = targetState;
                    }
                }
            }
            lastCheckedIndex = currentIndex;
        }

        // ── Filtering and Tools ──
        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string filter = txtSearch.Text.ToLower();
            FilteredEntries.Clear();
            lastCheckedIndex = -1; // Reset memory when filtering

            foreach (var item in AllEntries)
            {
                if (string.IsNullOrEmpty(filter) ||
                    item.OriginalName.ToLower().Contains(filter) ||
                    item.ViewType.ToLower().Contains(filter) ||
                    (item.Sheets != null && item.Sheets.ToLower().Contains(filter)))
                {
                    FilteredEntries.Add(item);
                }
            }
        }

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in FilteredEntries)
            {
                item.IsSelected = true;
            }
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in FilteredEntries)
            {
                item.IsSelected = false;
            }
            lastCheckedIndex = -1;
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in FilteredEntries)
            {
                item.NewName = item.OriginalName;
            }
        }

        // ── Counter Updates ──
        private void Entry_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(ViewAuditEntry.IsSelected) ||
                e.PropertyName == nameof(ViewAuditEntry.NameChanged))
            {
                UpdateStatus();
            }
        }

        private void UpdateStatus()
        {
            lblTotal.Text = $"Total Views: {AllEntries.Count}";
            lblSelected.Text = $"Selected: {AllEntries.Count(x => x.IsSelected)}";
            lblModified.Text = $"Names Modified: {AllEntries.Count(x => x.NameChanged)}";
        }
    }
}