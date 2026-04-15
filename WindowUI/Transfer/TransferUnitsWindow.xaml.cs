using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;

namespace HMVTools
{
    public partial class TransferUnitsWindow : Window
    {
        // ObservableCollection ensures the ListBox updates automatically
        private ObservableCollection<TargetDocEntry> _targets = new ObservableCollection<TargetDocEntry>();

        public TransferUnitsWindow(string sourceTitle, List<TargetDocEntry> openDocs)
        {
            InitializeComponent();
            txtSourceInfo.Text = $"Source Master:  {sourceTitle}";

            // Populate currently open documents
            foreach (var doc in openDocs)
            {
                _targets.Add(doc);
            }

            // Bind to the ListBox
            lstTargets.ItemsSource = _targets;
        }

        public List<TargetDocEntry> GetSelectedTargets()
        {
            return _targets.Where(t => t.IsSelected).ToList();
        }

        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Revit Files (*.rvt)|*.rvt",
                Multiselect = true,
                Title = "Select Target Revit Models"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string file in openFileDialog.FileNames)
                {
                    // Avoid adding duplicates
                    if (!_targets.Any(t => t.PathName.Equals(file, StringComparison.OrdinalIgnoreCase)))
                    {
                        _targets.Add(new TargetDocEntry
                        {
                            Title = System.IO.Path.GetFileName(file) + " (External)",
                            PathName = file,
                            OpenDoc = null,
                            IsOpenInRevit = false,
                            IsSelected = true
                        });
                    }
                }
            }
        }

        private void BtnTransfer_Click(object sender, RoutedEventArgs e)
        {
            if (GetSelectedTargets().Count == 0)
            {
                MessageBox.Show("Please select at least one target document.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
            Close();
        }
    }
}