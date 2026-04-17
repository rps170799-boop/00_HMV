using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;

namespace HMVTools
{
    public partial class FolderBrowserWindow : Window
    {
        private readonly ApsManager _aps;
        private readonly string _hubId;

        // Navigation state
        private enum BrowseLevel { Projects, TopFolders, SubFolders }
        private BrowseLevel _currentLevel = BrowseLevel.Projects;
        private List<(string Id, string Name)> _currentItems = new List<(string, string)>();

        // Selection history for "Back" navigation
        private readonly Stack<(BrowseLevel Level, List<(string Id, string Name)> Items, string Label, string Path)> _history
            = new Stack<(BrowseLevel, List<(string, string)>, string, string)>();

        // Selected values (output)
        public string SelectedProjectId { get; private set; }
        public string SelectedProjectName { get; private set; }
        public string SelectedFolderId { get; private set; }
        public string SelectedFolderDisplayPath { get; private set; }

        public FolderBrowserWindow(ApsManager aps, string hubId, string hubName)
        {
            InitializeComponent();
            _aps = aps;
            _hubId = hubId;

            BreadcrumbLabel.Text = hubName;
            PathLabel.Text = "Select a project";

            // Load projects
            LoadProjects();
        }

        private void LoadProjects()
        {
            try
            {
                var projects = System.Threading.Tasks.Task.Run(async () =>
                {
                    return await _aps.GetProjectsAsync(_hubId);
                }).GetAwaiter().GetResult();

                _currentItems = projects;
                _currentItems.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));
                _currentLevel = BrowseLevel.Projects;

                FolderList.Items.Clear();
                foreach (var p in _currentItems)
                    FolderList.Items.Add("📁  " + p.Name);

                PathLabel.Text = "Double-click a project to open it";
                UpdateBackButton();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading projects:\n{ex.Message}", "HMV Tools");
            }
        }

        private void LoadTopFolders(string projectId, string projectName)
        {
            try
            {
                var folders = System.Threading.Tasks.Task.Run(async () =>
                {
                    return await _aps.GetTopFoldersAsync(_hubId, projectId);
                }).GetAwaiter().GetResult();

                // Save current state for Back
                PushHistory();

                SelectedProjectId = projectId;
                SelectedProjectName = projectName;
                _currentItems = folders;
                _currentItems.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));
                _currentLevel = BrowseLevel.TopFolders;

                FolderList.Items.Clear();
                foreach (var f in _currentItems)
                    FolderList.Items.Add("📁  " + f.Name);

                BreadcrumbLabel.Text = $"{SelectedProjectName}";
                PathLabel.Text = "Double-click a folder to go inside, or click 'Search .rvt here'";
                SelectedFolderDisplayPath = SelectedProjectName;
                UpdateBackButton();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading folders:\n{ex.Message}", "HMV Tools");
            }
        }

        private void LoadSubFolders(string projectId, string folderId, string folderName)
        {
            try
            {
                var subFolders = System.Threading.Tasks.Task.Run(async () =>
                {
                    return await _aps.GetSubFoldersAsync(projectId, folderId);
                }).GetAwaiter().GetResult();

                if (subFolders.Count == 0)
                {
                    MessageBox.Show($"'{folderName}' has no subfolders.\nClick 'Search .rvt here' to scan it.",
                        "HMV Tools");
                    return;
                }

                // Save current state for Back
                PushHistory();

                SelectedFolderId = folderId;
                _currentItems = subFolders;
                _currentItems.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));
                _currentLevel = BrowseLevel.SubFolders;

                FolderList.Items.Clear();
                foreach (var f in _currentItems)
                    FolderList.Items.Add("📁  " + f.Name);

                SelectedFolderDisplayPath += " / " + folderName;
                BreadcrumbLabel.Text = SelectedFolderDisplayPath;
                PathLabel.Text = "Double-click to go deeper, or click 'Search .rvt here'";
                UpdateBackButton();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading subfolders:\n{ex.Message}", "HMV Tools");
            }
        }

        private void FolderList_DoubleClick(object sender, MouseButtonEventArgs e)
        {
            int idx = FolderList.SelectedIndex;
            if (idx < 0 || idx >= _currentItems.Count) return;

            var selected = _currentItems[idx];

            switch (_currentLevel)
            {
                case BrowseLevel.Projects:
                    LoadTopFolders(selected.Id, selected.Name);
                    break;

                case BrowseLevel.TopFolders:
                    SelectedFolderId = selected.Id;
                    LoadSubFolders(SelectedProjectId, selected.Id, selected.Name);
                    break;

                case BrowseLevel.SubFolders:
                    LoadSubFolders(SelectedProjectId, selected.Id, selected.Name);
                    break;
            }
        }

        private void SearchHere_Click(object sender, RoutedEventArgs e)
        {
            // If still at project level, user must drill into a folder first
            if (_currentLevel == BrowseLevel.Projects)
            {
                MessageBox.Show("Please select a project first by double-clicking it.",
                    "HMV Tools");
                return;
            }

            // If a folder is selected in the list, use that one
            int idx = FolderList.SelectedIndex;
            if (idx >= 0 && idx < _currentItems.Count)
            {
                SelectedFolderId = _currentItems[idx].Id;
                SelectedFolderDisplayPath += " / " + _currentItems[idx].Name;
            }

            // If no folder ID yet (user is at TopFolders level without selecting),
            // they need to select one
            if (string.IsNullOrEmpty(SelectedFolderId))
            {
                MessageBox.Show("Click a folder to select it, then click 'Search .rvt here'.",
                    "HMV Tools");
                return;
            }

            DialogResult = true;
            Close();
        }

        private void Back_Click(object sender, RoutedEventArgs e)
        {
            if (_history.Count == 0) return;

            var prev = _history.Pop();
            _currentLevel = prev.Level;
            _currentItems = prev.Items;

            FolderList.Items.Clear();
            string icon = "📁  ";
            foreach (var item in _currentItems)
                FolderList.Items.Add(icon + item.Name);

            BreadcrumbLabel.Text = prev.Label;
            SelectedFolderDisplayPath = prev.Path;

            if (_currentLevel == BrowseLevel.Projects)
            {
                PathLabel.Text = "Double-click a project to open it";
                SelectedFolderId = null;
            }
            else
            {
                PathLabel.Text = "Double-click a folder to go inside, or click 'Search .rvt here'";
            }

            UpdateBackButton();
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

        private void PushHistory()
        {
            _history.Push((_currentLevel, new List<(string, string)>(_currentItems),
                BreadcrumbLabel.Text, SelectedFolderDisplayPath ?? ""));
        }

        private void UpdateBackButton()
        {
            BackButton.IsEnabled = _history.Count > 0;
        }
    }
}