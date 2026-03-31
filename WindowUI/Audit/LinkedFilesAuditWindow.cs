using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Documents;
using System.Text.RegularExpressions; // <-- Added for Regex standardization

using OfficeOpenXml;

namespace HMVTools
{
    public class LinkedFilesAuditWindow : Window
    {
        private List<LinkAuditData> _auditData;
        private ListView dataListView;
        private TextBox searchBox;
        private string _mainProjectName;

        private static readonly Color BluePrimary = Color.FromRgb(0, 120, 212);
        private static readonly Color GrayBg = Color.FromRgb(240, 240, 243);
        private static readonly Color DarkText = Color.FromRgb(30, 30, 30);
        private static readonly Color MutedText = Color.FromRgb(120, 120, 130);
        private static readonly Color BorderColor = Color.FromRgb(200, 200, 210);
        private static readonly Color WindowBg = Color.FromRgb(245, 245, 248);

        public LinkedFilesAuditWindow(List<LinkAuditData> auditData, string projectName)
        {
            _mainProjectName = projectName;

            // --- 0. STANDARDIZE THE SURVEY POINT DATA ---
            // Clean up units, spaces, and commas before doing ANY sorting or analysis
            foreach (var item in auditData)
            {
                item.SurveyPoint = StandardizeSurveyPoint(item.SurveyPoint);
            }

            // Sort first by Building Description, then alphabetically by Link Name
            _auditData = auditData
                .OrderBy(x => x.BuildingDescription)
                .ThenBy(x => x.LinkName)
                .ToList();

            // --- 1. FREQUENCY ANALYSIS FOR SURVEY POINT ---
            string mostFrequentSurveyPoint = _auditData
                .Where(x => !string.IsNullOrWhiteSpace(x.SurveyPoint) && x.SurveyPoint != "N/A" && x.SurveyPoint != "UNLOADED")
                .GroupBy(x => x.SurveyPoint)
                .OrderByDescending(g => g.Count())
                .Select(g => g.Key)
                .FirstOrDefault();

            Title = "HMV Tools - Linked Files Audit";
            Width = 1700;
            Height = 900;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(WindowBg);

            var mainGrid = new Grid();
            mainGrid.Margin = new Thickness(20);
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // ── Row 0: Header Area ────────────────
            var headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            headerGrid.Margin = new Thickness(0, 0, 0, 15);

            var titlePanel = new StackPanel { Orientation = Orientation.Vertical };
            titlePanel.Children.Add(new TextBlock { Text = "Linked Files Coordinate Audit", FontSize = 18, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(DarkText), Margin = new Thickness(0, 0, 0, 4) });
            titlePanel.Children.Add(new TextBlock { Text = "Review GIS, descriptions, and coordinate data for all loaded Revit links.", FontSize = 12, Foreground = new SolidColorBrush(MutedText) });

            Grid.SetColumn(titlePanel, 0);
            headerGrid.Children.Add(titlePanel);

            var searchPanel = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
            searchPanel.Children.Add(new TextBlock { Text = "Search:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 8, 0), Foreground = new SolidColorBrush(DarkText) });

            searchBox = new TextBox { Width = 200, Height = 28, VerticalContentAlignment = VerticalAlignment.Center, Padding = new Thickness(5, 0, 5, 0) };
            searchBox.TextChanged += SearchBox_TextChanged;
            searchPanel.Children.Add(searchBox);

            Grid.SetColumn(searchPanel, 1);
            headerGrid.Children.Add(searchPanel);

            Grid.SetRow(headerGrid, 0);
            mainGrid.Children.Add(headerGrid);

            // ── Row 1: Data Table ───────────────────────
            var listBorder = new System.Windows.Controls.Border { CornerRadius = new CornerRadius(8), BorderBrush = new SolidColorBrush(BorderColor), BorderThickness = new Thickness(1), Background = Brushes.White, Margin = new Thickness(0, 0, 0, 15) };

            dataListView = new ListView
            {
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                FontSize = 10,
                Margin = new Thickness(4),
                SelectionMode = SelectionMode.Single
            };

            ScrollViewer.SetCanContentScroll(dataListView, false);

            var itemStyle = new Style(typeof(ListViewItem));

            var backgroundBinding = new Binding("BuildingName");
            backgroundBinding.Converter = new BuildingNameColorConverter();
            itemStyle.Setters.Add(new Setter(ListViewItem.BackgroundProperty, backgroundBinding));

            var template = new ControlTemplate(typeof(ListViewItem));

            var borderFactory = new FrameworkElementFactory(typeof(System.Windows.Controls.Border));
            borderFactory.Name = "RowBorder";
            borderFactory.SetBinding(System.Windows.Controls.Border.BackgroundProperty, new Binding("Background") { RelativeSource = RelativeSource.TemplatedParent });
            borderFactory.SetValue(System.Windows.Controls.Border.BorderBrushProperty, Brushes.Transparent);
            borderFactory.SetValue(System.Windows.Controls.Border.BorderThicknessProperty, new Thickness(1));
            borderFactory.SetValue(System.Windows.Controls.Border.PaddingProperty, new Thickness(0, 4, 0, 4));

            var presenterFactory = new FrameworkElementFactory(typeof(GridViewRowPresenter));
            presenterFactory.SetValue(GridViewRowPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);

            borderFactory.AppendChild(presenterFactory);
            template.VisualTree = borderFactory;

            var selectedTrigger = new Trigger { Property = ListViewItem.IsSelectedProperty, Value = true };
            selectedTrigger.Setters.Add(new Setter(System.Windows.Controls.Border.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(150, 150, 150)), "RowBorder"));
            selectedTrigger.Setters.Add(new Setter(TextElement.FontWeightProperty, FontWeights.Bold));
            template.Triggers.Add(selectedTrigger);

            itemStyle.Setters.Add(new Setter(ListViewItem.TemplateProperty, template));
            dataListView.ItemContainerStyle = itemStyle;

            // ── Columns ───────────────────────
            var surveyPointColorConverter = new FrequencyTextColorConverter(mostFrequentSurveyPoint);

            var gridView = new GridView();
            gridView.Columns.Add(CreateColumn("Revit Link Name", "LinkName", 300));
            gridView.Columns.Add(CreateColumn("Building Desc.", "BuildingDescription", 150));
            gridView.Columns.Add(CreateColumn("Building Name", "BuildingName", 150));
            gridView.Columns.Add(CreateColumn("GIS Code", "GisCode", 100));
            gridView.Columns.Add(CreateColumn("Latitude", "Latitude", 100));
            gridView.Columns.Add(CreateColumn("Longitude", "Longitude", 100));
            gridView.Columns.Add(CreateColumn("True North", "TrueNorthAngle", 100));
            gridView.Columns.Add(CreateColumn("Site Name", "SiteName", 120));
            gridView.Columns.Add(CreateColumn("Project Base Point", "ProjectBasePoint", 220));
            gridView.Columns.Add(CreateColumn("Survey Point", "SurveyPoint", 220, surveyPointColorConverter));

            dataListView.View = gridView;
            dataListView.ItemsSource = _auditData;

            listBorder.Child = dataListView;
            Grid.SetRow(listBorder, 1);
            mainGrid.Children.Add(listBorder);

            // ── Row 2: Buttons ─────────────────────────────────────
            var buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };

            var closeBtn = CreateButton("Close", GrayBg, Color.FromRgb(60, 60, 60));
            closeBtn.Click += (s, e) => { DialogResult = false; Close(); };
            closeBtn.Margin = new Thickness(0, 0, 8, 0);
            closeBtn.Width = 90;

            var exportBtn = CreateButton("Export Excel", BluePrimary, Color.FromRgb(255, 255, 255));
            exportBtn.Click += (s, e) => ExportToExcel();
            exportBtn.Width = 120;

            buttonPanel.Children.Add(closeBtn);
            buttonPanel.Children.Add(exportBtn);
            Grid.SetRow(buttonPanel, 2);
            mainGrid.Children.Add(buttonPanel);

            Content = mainGrid;
        }

        // --- NEW METHOD: Cleans and standardizes the Survey Point string ---
        // --- NEW METHOD: Cleans and standardizes the Survey Point string ---
        private string StandardizeSurveyPoint(string input)
        {
            if (string.IsNullOrWhiteSpace(input) || input == "N/A" || input == "UNLOADED")
                return input;

            try
            {
                // This Regex strictly hunts for the numbers next to N/S, E/W, and Elev. 
                var match = Regex.Match(input, @"N/S:\s*([-\d.,]+)[^E]*E/W:\s*([-\d.,]+)[^E]*Elev:\s*([-\d.,]+)");

                if (match.Success)
                {
                    // Clean up any trailing commas or spaces that might have been grabbed
                    string rawNs = match.Groups[1].Value.TrimEnd(',', ' ');
                    string rawEw = match.Groups[2].Value.TrimEnd(',', ' ');
                    string rawElev = match.Groups[3].Value.TrimEnd(',', ' ');

                    // Parse the numbers, round them, and force a period (.) for the decimal
                    string ns = FormatCoordinate(rawNs);
                    string ew = FormatCoordinate(rawEw);
                    string elev = FormatCoordinate(rawElev);

                    // Rebuild the string perfectly every time
                    return $"N/S: {ns}, E/W: {ew}, Elev: {elev}";
                }
            }
            catch
            {
                // Silently fall back to simple replacement if something goes completely wrong
            }

            // Fallback: Just strip out the 'm' and double spaces
            return input.Replace(" m", "").Replace(" mm", "").Replace("  ", " ").Trim();
        }

        // --- NEW HELPER METHOD: Standardizes commas, periods, and decimal places ---
        private string FormatCoordinate(string value)
        {
            // Force replace commas with periods (assuming no thousands separators are used in raw Revit output)
            string normalizedString = value.Replace(",", ".");

            // Try to parse it into an actual number
            if (double.TryParse(normalizedString, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedValue))
            {
                // Format to exactly 2 decimal places ("F2") to match your standard, 
                // and use InvariantCulture to ALWAYS use a period (.) for decimals.
                // Note: If you want 3 decimal places, change "F2" to "F3".
                return parsedValue.ToString("F4", System.Globalization.CultureInfo.InvariantCulture);
            }

            return value; // If parsing fails for some reason, return the string as-is
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string filterText = searchBox.Text.ToLower();
            if (string.IsNullOrWhiteSpace(filterText))
            {
                dataListView.ItemsSource = _auditData;
            }
            else
            {
                dataListView.ItemsSource = _auditData.Where(x => x.LinkName.ToLower().Contains(filterText)).ToList();
            }
        }

        private GridViewColumn CreateColumn(string header, string bindingPath, double width, IValueConverter textColorConverter = null)
        {
            var column = new GridViewColumn { Header = header, Width = width };

            var template = new DataTemplate();
            var textBlockFactory = new FrameworkElementFactory(typeof(TextBlock));

            textBlockFactory.SetBinding(TextBlock.TextProperty, new Binding(bindingPath));
            textBlockFactory.SetValue(TextBlock.TextWrappingProperty, TextWrapping.Wrap);
            textBlockFactory.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
            textBlockFactory.SetValue(TextBlock.MarginProperty, new Thickness(2, 2, 8, 2));

            if (textColorConverter != null)
            {
                var colorBinding = new Binding(bindingPath);
                colorBinding.Converter = textColorConverter;
                textBlockFactory.SetBinding(TextBlock.ForegroundProperty, colorBinding);
                textBlockFactory.SetValue(TextBlock.FontWeightProperty, FontWeights.SemiBold);
            }

            template.VisualTree = textBlockFactory;
            column.CellTemplate = template;

            return column;
        }

        private void ExportToExcel()
        {
            var currentData = dataListView.ItemsSource as List<LinkAuditData>;
            if (currentData == null || currentData.Count == 0) return;

            Microsoft.Win32.SaveFileDialog sfd = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = $"{_mainProjectName}_LinkedFiles_Audit.xlsx",
                Title = "Save Audit Report"
            };

            if (sfd.ShowDialog() == true)
            {
                try
                {
                    ExcelPackage.License.SetNonCommercialOrganization("HMV Ingenieros");

                    using (var package = new ExcelPackage())
                    {
                        var worksheet = package.Workbook.Worksheets.Add("Linked Files Audit");

                        worksheet.Cells[1, 1].Value = "Host File Name";
                        worksheet.Cells[1, 2].Value = "Revit Link Name";
                        worksheet.Cells[1, 3].Value = "Building Description";
                        worksheet.Cells[1, 4].Value = "Building Name";
                        worksheet.Cells[1, 5].Value = "GIS Coordinate System";
                        worksheet.Cells[1, 6].Value = "Latitude";
                        worksheet.Cells[1, 7].Value = "Longitude";
                        worksheet.Cells[1, 8].Value = "Site Name";
                        worksheet.Cells[1, 9].Value = "Angle to True North";
                        worksheet.Cells[1, 10].Value = "Project Base Point";
                        worksheet.Cells[1, 11].Value = "Survey Point";

                        var headerRange = worksheet.Cells["A1:K1"];
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Font.Color.SetColor(System.Drawing.Color.White);
                        headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkBlue);
                        headerRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                        int currentRow = 2;
                        foreach (var item in currentData)
                        {
                            worksheet.Cells[currentRow, 1].Value = _mainProjectName;
                            worksheet.Cells[currentRow, 2].Value = item.LinkName;
                            worksheet.Cells[currentRow, 3].Value = item.BuildingDescription;
                            worksheet.Cells[currentRow, 4].Value = item.BuildingName;
                            worksheet.Cells[currentRow, 5].Value = item.GisCode;
                            worksheet.Cells[currentRow, 6].Value = item.Latitude;
                            worksheet.Cells[currentRow, 7].Value = item.Longitude;
                            worksheet.Cells[currentRow, 8].Value = item.SiteName;
                            worksheet.Cells[currentRow, 9].Value = item.TrueNorthAngle;
                            worksheet.Cells[currentRow, 10].Value = item.ProjectBasePoint;
                            worksheet.Cells[currentRow, 11].Value = item.SurveyPoint;
                            currentRow++;
                        }

                        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                        worksheet.Column(1).Width = Math.Min(worksheet.Column(1).Width, 45);
                        worksheet.Column(2).Width = Math.Min(worksheet.Column(2).Width, 35);
                        worksheet.Column(9).Width = Math.Min(worksheet.Column(9).Width, 35);
                        worksheet.Column(10).Width = Math.Min(worksheet.Column(10).Width, 35);

                        var dataRange = worksheet.Cells[2, 1, currentRow - 1, 11];
                        dataRange.Style.WrapText = true;
                        dataRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                        package.SaveAs(new FileInfo(sfd.FileName));
                    }

                    System.Windows.MessageBox.Show("Audit report successfully exported to Excel!", "HMV Tools", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"Error saving Excel file:\n\n{ex.Message}", "HMV Tools", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                }
            }
        }

        private Button CreateButton(string text, Color bgColor, Color fgColor)
        {
            var btn = new Button { Content = text, Height = 36, FontSize = 13, Foreground = new SolidColorBrush(fgColor), Background = new SolidColorBrush(bgColor), BorderThickness = new Thickness(0), Cursor = Cursors.Hand };
            btn.Template = GetRoundButtonTemplate(bgColor);
            return btn;
        }

        private ControlTemplate GetRoundButtonTemplate(Color bgColor)
        {
            var template = new ControlTemplate(typeof(Button));
            var border = new FrameworkElementFactory(typeof(System.Windows.Controls.Border));
            border.SetValue(System.Windows.Controls.Border.CornerRadiusProperty, new CornerRadius(6));
            border.SetValue(System.Windows.Controls.Border.BackgroundProperty, new SolidColorBrush(bgColor));
            border.SetValue(System.Windows.Controls.Border.PaddingProperty, new Thickness(14, 6, 14, 6));

            var content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            content.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);

            border.AppendChild(content);
            template.VisualTree = border;
            return template;
        }
    }

    public class FrequencyTextColorConverter : IValueConverter
    {
        private string _standardValue;
        private SolidColorBrush _highlightColor = new SolidColorBrush(Color.FromRgb(160, 0, 0)); // Dark Red
        private SolidColorBrush _normalColor = new SolidColorBrush(Color.FromRgb(30, 30, 30));   // Standard Text

        public FrequencyTextColorConverter(string standardValue)
        {
            _standardValue = standardValue;
        }

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            string cellValue = value as string;

            if (!string.IsNullOrEmpty(cellValue) && cellValue == _standardValue)
            {
                return _highlightColor;
            }

            return _normalColor;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class BuildingNameColorConverter : IValueConverter
    {
        private Dictionary<string, SolidColorBrush> _colorMap = new Dictionary<string, SolidColorBrush>();
        private int _colorIndex = 0;

        private List<Color> _palette = new List<Color>
        {
            Color.FromRgb(227, 242, 253), // Light Blue
            Color.FromRgb(232, 245, 233), // Light Green
            Color.FromRgb(255, 243, 224), // Light Orange
            Color.FromRgb(243, 229, 245), // Light Purple
            Color.FromRgb(255, 235, 238), // Light Red
            Color.FromRgb(255, 253, 231), // Light Yellow
            Color.FromRgb(224, 242, 241)  // Light Teal
        };

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            string buildingName = value as string;

            if (string.IsNullOrEmpty(buildingName) || buildingName == "N/A" || buildingName == "UNLOADED")
            {
                return Brushes.Transparent;
            }

            if (!_colorMap.ContainsKey(buildingName))
            {
                _colorMap[buildingName] = new SolidColorBrush(_palette[_colorIndex % _palette.Count]);
                _colorIndex++;
            }

            return _colorMap[buildingName];
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}