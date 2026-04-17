using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using OfficeOpenXml;

namespace HMVTools
{
    public partial class LinkedFilesAuditWindow : Window
    {
        private List<LinkAuditData> _auditData;
        private string _mainProjectName;
        private string _mostFrequentSurveyPoint;
        private GridViewColumnHeader _lastHeaderClicked = null;
        private System.ComponentModel.ListSortDirection _lastDirection = System.ComponentModel.ListSortDirection.Ascending;

        public LinkedFilesAuditWindow(List<LinkAuditData> auditData, string projectName)
        {
            InitializeComponent();

            _mainProjectName = projectName;
            auditData = auditData ?? new List<LinkAuditData>();

            // 0. Standardize survey points BEFORE sorting / frequency analysis
            foreach (var item in auditData)
            {
                item.SurveyPoint = StandardizeSurveyPoint(item.SurveyPoint);
            }

            // Sort: Force Host Model to bottom, then BuildingDescription, then LinkName
            _auditData = auditData
                .OrderBy(x => x.LinkName != null && x.LinkName.StartsWith("--- HOST") ? 1 : 0)
                .ThenBy(x => x.BuildingDescription)
                .ThenBy(x => x.LinkName)
                .ToList();

            // 1. Frequency analysis — find the "standard" Survey Point
            _mostFrequentSurveyPoint = _auditData
                .Where(x => !string.IsNullOrWhiteSpace(x.SurveyPoint)
                            && x.SurveyPoint != "N/A"
                            && x.SurveyPoint != "UNLOADED")
                .GroupBy(x => x.SurveyPoint)
                .OrderByDescending(g => g.Count())
                .Select(g => g.Key)
                .FirstOrDefault();

            SetupListViewColumns();
            dataListView.ItemsSource = _auditData;
        }

        // ── Window chrome ─────────────────────────────────────────
        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        // ── Columns (matches LinkAuditData properties EXACTLY) ────
        private void SetupListViewColumns()
        {
            var gridView = new GridView();

            var surveyPointConverter = new FrequencyTextColorConverter(_mostFrequentSurveyPoint);

            gridView.Columns.Add(CreateColumn("Revit Link Name", "LinkName", 260));
            gridView.Columns.Add(CreateColumn("Building Name", "BuildingName", 140));
            gridView.Columns.Add(CreateColumn("Building Desc.", "BuildingDescription", 140));
            gridView.Columns.Add(CreateColumn("Site Name", "SiteName", 120));
            gridView.Columns.Add(CreateColumn("GIS Code", "GisCode", 100));
            gridView.Columns.Add(CreateColumn("Latitude", "Latitude", 90));
            gridView.Columns.Add(CreateColumn("Longitude", "Longitude", 90));
            gridView.Columns.Add(CreateColumn("True North", "TrueNorthAngle", 90));
            gridView.Columns.Add(CreateColumn("Project Base Point", "ProjectBasePoint", 210));
            gridView.Columns.Add(CreateColumn("Survey Point", "SurveyPoint", 210, surveyPointConverter));

            dataListView.View = gridView;

            // Row background color = Building Name group color
            var itemStyle = new Style(typeof(ListViewItem), dataListView.ItemContainerStyle);
            itemStyle.Setters.Add(new Setter(
                ListViewItem.BackgroundProperty,
                new Binding("BuildingName") { Converter = new BuildingNameColorConverter() }
            ));
            dataListView.ItemContainerStyle = itemStyle;
        }

        private GridViewColumn CreateColumn(string header, string bindingPath, double width, IValueConverter textColorConverter = null)
        {
            var column = new GridViewColumn { Header = header, Width = width };

            var template = new DataTemplate();
            var tb = new FrameworkElementFactory(typeof(TextBlock));
            tb.SetBinding(TextBlock.TextProperty, new Binding(bindingPath));
            tb.SetValue(TextBlock.TextWrappingProperty, TextWrapping.Wrap);
            tb.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
            tb.SetValue(TextBlock.MarginProperty, new Thickness(2, 2, 8, 2));

            if (textColorConverter != null)
            {
                var colorBinding = new Binding(bindingPath) { Converter = textColorConverter };
                tb.SetBinding(TextBlock.ForegroundProperty, colorBinding);
                tb.SetValue(TextBlock.FontWeightProperty, FontWeights.SemiBold);
            }

            template.VisualTree = tb;
            column.CellTemplate = template;
            return column;
        }

        // ── Search (original: filters by LinkName only) ───────────
        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string query = searchBox.Text?.ToLower() ?? "";
            if (string.IsNullOrWhiteSpace(query))
            {
                dataListView.ItemsSource = _auditData;
            }
            else
            {
                dataListView.ItemsSource = _auditData
                    .Where(x => x.LinkName != null && x.LinkName.ToLower().Contains(query))
                    .ToList();
            }
        }
        // ── Column Header Click Sorting ──────────────────────────
        private void GridViewColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            var headerClicked = e.OriginalSource as GridViewColumnHeader;
            if (headerClicked == null || headerClicked.Role == GridViewColumnHeaderRole.Padding) return;

            string header = headerClicked.Column.Header as string;
            if (string.IsNullOrEmpty(header)) return;

            // Match header to property name
            string sortBy = "";
            switch (header)
            {
                case "Revit Link Name": sortBy = "LinkName"; break;
                case "Building Name": sortBy = "BuildingName"; break;
                case "Building Desc.": sortBy = "BuildingDescription"; break;
                case "Site Name": sortBy = "SiteName"; break;
                case "GIS Code": sortBy = "GisCode"; break;
                case "Latitude": sortBy = "Latitude"; break;
                case "Longitude": sortBy = "Longitude"; break;
                case "True North": sortBy = "TrueNorthAngle"; break;
                case "Project Base Point": sortBy = "ProjectBasePoint"; break;
                case "Survey Point": sortBy = "SurveyPoint"; break;
            }

            if (string.IsNullOrEmpty(sortBy)) return;

            // Toggle sorting direction
            System.ComponentModel.ListSortDirection direction;
            if (headerClicked != _lastHeaderClicked)
            {
                direction = System.ComponentModel.ListSortDirection.Ascending;
            }
            else
            {
                direction = _lastDirection == System.ComponentModel.ListSortDirection.Ascending
                    ? System.ComponentModel.ListSortDirection.Descending
                    : System.ComponentModel.ListSortDirection.Ascending;
            }

            SortList(sortBy, direction);

            _lastHeaderClicked = headerClicked;
            _lastDirection = direction;
        }

        private void SortList(string sortBy, System.ComponentModel.ListSortDirection direction)
        {
            var propInfo = typeof(LinkAuditData).GetProperty(sortBy);
            if (propInfo == null) return;

            // Sort data while forcing Host model to remain at the bottom
            if (direction == System.ComponentModel.ListSortDirection.Ascending)
            {
                _auditData = _auditData
                    .OrderBy(x => x.LinkName != null && x.LinkName.StartsWith("--- HOST") ? 1 : 0)
                    .ThenBy(x => propInfo.GetValue(x, null))
                    .ToList();
            }
            else
            {
                _auditData = _auditData
                    .OrderBy(x => x.LinkName != null && x.LinkName.StartsWith("--- HOST") ? 1 : 0)
                    .ThenByDescending(x => propInfo.GetValue(x, null))
                    .ToList();
            }

            // Re-apply search filter to update the view
            SearchBox_TextChanged(null, null);
        }
        // ── Survey Point standardization ──────────────────────────
        private string StandardizeSurveyPoint(string input)
        {
            if (string.IsNullOrWhiteSpace(input) || input == "N/A" || input == "UNLOADED")
                return input;

            try
            {
                var match = Regex.Match(input, @"N/S:\s*([-\d.,]+)[^E]*E/W:\s*([-\d.,]+)[^E]*Elev:\s*([-\d.,]+)");
                if (match.Success)
                {
                    string ns = FormatCoordinate(match.Groups[1].Value.TrimEnd(',', ' '));
                    string ew = FormatCoordinate(match.Groups[2].Value.TrimEnd(',', ' '));
                    string elev = FormatCoordinate(match.Groups[3].Value.TrimEnd(',', ' '));
                    return $"N/S: {ns}, E/W: {ew}, Elev: {elev}";
                }
            }
            catch { /* fall through */ }

            return input.Replace(" m", "").Replace(" mm", "").Replace("  ", " ").Trim();
        }

        private string FormatCoordinate(string value)
        {
            string normalized = value.Replace(",", ".");
            if (double.TryParse(normalized,
                                System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture,
                                out double parsed))
            {
                return parsed.ToString("F10", System.Globalization.CultureInfo.InvariantCulture);
            }
            return value;
        }

        // ── Excel export (original 11-column layout) ──────────────
        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            var currentData = dataListView.ItemsSource as IEnumerable<LinkAuditData>;
            var list = currentData?.ToList() ?? new List<LinkAuditData>();

            if (list.Count == 0)
            {
                MessageBox.Show("No data available to export.", "HMV Tools",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var sfd = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = $"{_mainProjectName}_LinkedFiles_Audit.xlsx",
                Title = "Save Audit Report"
            };

            if (sfd.ShowDialog() != true) return;

            try
            {
                ExcelPackage.License.SetNonCommercialOrganization("HMV Ingenieros");

                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Linked Files Audit");

                    
                    ws.Cells[1, 1].Value = "Revit Link Name";
                    ws.Cells[1, 2].Value = "Building Name";
                    ws.Cells[1, 3].Value = "Building Description";
                    ws.Cells[1, 4].Value = "Site Name";
                    ws.Cells[1, 5].Value = "GIS Coordinate System";
                    ws.Cells[1, 6].Value = "Latitude";
                    ws.Cells[1, 7].Value = "Longitude";
                    ws.Cells[1, 8].Value = "Angle to True North";
                    ws.Cells[1, 9].Value = "Project Base Point";
                    ws.Cells[1, 10].Value = "Survey Point";

                    var headerRange = ws.Cells["A1:J1"];
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 73, 130));
                    headerRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                    int row = 2;
                    foreach (var item in list)
                    {
                       // ws.Cells[row, 1].Value = _mainProjectName;
                        ws.Cells[row, 1].Value = item.LinkName;
                        ws.Cells[row, 2].Value = item.BuildingName;
                        ws.Cells[row, 3].Value = item.BuildingDescription;
                        ws.Cells[row, 4].Value = item.SiteName;
                        ws.Cells[row, 5].Value = item.GisCode;
                        ws.Cells[row, 6].Value = item.Latitude;
                        ws.Cells[row, 7].Value = item.Longitude;
                        ws.Cells[row, 8].Value = item.TrueNorthAngle;
                        ws.Cells[row, 9].Value = item.ProjectBasePoint;
                        ws.Cells[row, 10].Value = item.SurveyPoint;
                        row++;
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    ws.Column(1).Width = Math.Min(ws.Column(1).Width, 45);
                    ws.Column(2).Width = Math.Min(ws.Column(2).Width, 35);
                    ws.Column(9).Width = Math.Min(ws.Column(9).Width, 35);
                    ws.Column(10).Width = Math.Min(ws.Column(10).Width, 35);

                    var dataRange = ws.Cells[2, 1, row - 1, 11];
                    dataRange.Style.WrapText = true;
                    dataRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                    package.SaveAs(new FileInfo(sfd.FileName));
                }

                MessageBox.Show("Audit report successfully exported to Excel!", "HMV Tools",
                                MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving Excel file:\n\n{ex.Message}", "HMV Tools",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    // ── Converters ────────────────────────────────────────────────
    public class FrequencyTextColorConverter : IValueConverter
    {
        private readonly string _standardValue;
        private readonly SolidColorBrush _highlightColor = new SolidColorBrush(Color.FromRgb(160, 0, 0));
        private readonly SolidColorBrush _normalColor = new SolidColorBrush(Color.FromRgb(30, 30, 30));

        public FrequencyTextColorConverter(string standardValue) { _standardValue = standardValue; }

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            string cell = value as string;
            if (!string.IsNullOrEmpty(cell) && cell == _standardValue) return _normalColor;
            return _highlightColor; // anything different from the standard gets flagged
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            => throw new NotImplementedException();
    }

    public class BuildingNameColorConverter : IValueConverter
    {
        private readonly Dictionary<string, SolidColorBrush> _colorMap =
            new Dictionary<string, SolidColorBrush>(StringComparer.OrdinalIgnoreCase);
        private int _colorIndex = 0;
        private readonly object _lock = new object();

        private readonly List<Color> _palette = new List<Color>
        {
            Color.FromRgb(227, 242, 253),
            Color.FromRgb(232, 245, 233),
            Color.FromRgb(255, 243, 224),
            Color.FromRgb(243, 229, 245),
            Color.FromRgb(255, 235, 238),
            Color.FromRgb(255, 253, 231),
            Color.FromRgb(224, 242, 241)
        };

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            string name = value as string;
            if (string.IsNullOrWhiteSpace(name)
                || name.Equals("N/A", StringComparison.OrdinalIgnoreCase)
                || name.Equals("UNLOADED", StringComparison.OrdinalIgnoreCase))
                return Brushes.Transparent;

            lock (_lock)
            {
                if (!_colorMap.TryGetValue(name, out SolidColorBrush brush))
                {
                    brush = new SolidColorBrush(_palette[_colorIndex % _palette.Count]);
                    _colorMap[name] = brush;
                    _colorIndex++;
                }
                return brush;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            => null;
    }
}