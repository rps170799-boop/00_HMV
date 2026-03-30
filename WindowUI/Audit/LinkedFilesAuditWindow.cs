using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32; 
using ClosedXML.Excel;

namespace HMVTools
{
    public class LinkedFilesAuditWindow : Window
    {
        // Data
        private List<LinkAuditData> _auditData;

        // Controls
        private ListView dataListView;

        // Colors (Matched to PipeAnnotationWindow)
        private static readonly Color BluePrimary = Color.FromRgb(0, 120, 212);
        private static readonly Color GrayBg = Color.FromRgb(240, 240, 243);
        private static readonly Color DarkText = Color.FromRgb(30, 30, 30);
        private static readonly Color MutedText = Color.FromRgb(120, 120, 130);
        private static readonly Color BorderColor = Color.FromRgb(200, 200, 210);
        private static readonly Color WindowBg = Color.FromRgb(245, 245, 248);

        public LinkedFilesAuditWindow(List<LinkAuditData> auditData)
        {
            _auditData = auditData;

            Title = "HMV Tools - Linked Files Audit";
            Width = 850;  // Wider to accommodate 6 columns of data
            Height = 500;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(WindowBg);

            var mainGrid = new Grid();
            mainGrid.Margin = new Thickness(20);
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 0 Title
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 1 Subtitle
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // 2 Table
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // 3 Buttons

            // ── Row 0: Title ───────────────────────────────────────
            var title = new TextBlock
            {
                Text = "Linked Files Coordinate Audit",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 6)
            };
            Grid.SetRow(title, 0);
            mainGrid.Children.Add(title);

            // ── Row 1: Subtitle ────────────────────────────────────
            var subtitleText = new TextBlock
            {
                Text = "Review the GIS and coordinate data for all loaded Revit links in the current model.",
                FontSize = 12,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(0, 0, 0, 15)
            };
            Grid.SetRow(subtitleText, 1);
            mainGrid.Children.Add(subtitleText);

            // ── Row 2: Data Table (ListView) ───────────────────────
            var listBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 15)
            };

            dataListView = new ListView
            {
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                FontSize = 13,
                Margin = new Thickness(4),
                SelectionMode = SelectionMode.Single
            };

            // Define GridView columns for the tabular data
            var gridView = new GridView();
            gridView.Columns.Add(CreateColumn("Link Name", "LinkName", 200));
            gridView.Columns.Add(CreateColumn("GIS Code", "GisCode", 120));
            gridView.Columns.Add(CreateColumn("Latitude", "Latitude", 100));
            gridView.Columns.Add(CreateColumn("Longitude", "Longitude", 100));
            gridView.Columns.Add(CreateColumn("Site Name", "SiteName", 150));
            gridView.Columns.Add(CreateColumn("True North", "TrueNorthAngle", 90));

            dataListView.View = gridView;
            dataListView.ItemsSource = _auditData;

            listBorder.Child = dataListView;
            Grid.SetRow(listBorder, 2);
            mainGrid.Children.Add(listBorder);

            // ── Row 3: Buttons ─────────────────────────────────────
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var closeBtn = CreateButton("Close", GrayBg, Color.FromRgb(60, 60, 60));
            closeBtn.Click += (s, e) => { DialogResult = false; Close(); };
            closeBtn.Margin = new Thickness(0, 0, 8, 0);
            closeBtn.Width = 90;

            var exportBtn = CreateButton("Export Excel", BluePrimary, Color.FromRgb(255, 255, 255));
            exportBtn.Click += (s, e) => ExportToExcel();
            exportBtn.Width = 120;

            buttonPanel.Children.Add(closeBtn);
            buttonPanel.Children.Add(exportBtn);
            Grid.SetRow(buttonPanel, 3);
            mainGrid.Children.Add(buttonPanel);

            Content = mainGrid;
        }

        // ── Helper: Create GridView Column ─────────────────────────
        private GridViewColumn CreateColumn(string header, string bindingPath, double width)
        {
            return new GridViewColumn
            {
                Header = header,
                DisplayMemberBinding = new Binding(bindingPath),
                Width = width
            };
        }

        // ── Export Logic ───────────────────────────────────────────
        private void ExportToExcel()
        {
            if (_auditData == null || _auditData.Count == 0)
            {
                MessageBox.Show("No data available to export.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = "LinkedFiles_Audit_Report.xlsx",
                Title = "Save Audit Report"
            };

            if (sfd.ShowDialog() == true)
            {
                try
                {
                    // Create a new Excel Workbook using ClosedXML
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Linked Files Audit");

                        // 1. Add Headers
                        worksheet.Cell(1, 1).Value = "Link Name";
                        worksheet.Cell(1, 2).Value = "GIS Coordinate System";
                        worksheet.Cell(1, 3).Value = "Latitude";
                        worksheet.Cell(1, 4).Value = "Longitude";
                        worksheet.Cell(1, 5).Value = "Site Name";
                        worksheet.Cell(1, 6).Value = "Angle to True North";

                        // Format Headers (Bold and slight background color)
                        var headerRange = worksheet.Range("A1:F1");
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

                        // 2. Add Data
                        int currentRow = 2;
                        foreach (var item in _auditData)
                        {
                            worksheet.Cell(currentRow, 1).Value = item.LinkName;
                            worksheet.Cell(currentRow, 2).Value = item.GisCode;
                            worksheet.Cell(currentRow, 3).Value = item.Latitude;
                            worksheet.Cell(currentRow, 4).Value = item.Longitude;
                            worksheet.Cell(currentRow, 5).Value = item.SiteName;
                            worksheet.Cell(currentRow, 6).Value = item.TrueNorthAngle;

                            currentRow++;
                        }

                        // Auto-fit columns for better readability
                        worksheet.Columns().AdjustToContents();

                        // 3. Save the file
                        workbook.SaveAs(sfd.FileName);
                    }

                    MessageBox.Show("Audit report successfully exported to Excel!", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error saving Excel file. Ensure the file isn't currently open.\n\n{ex.Message}", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // ── UI helpers (Matched from your reference) ───────────────

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