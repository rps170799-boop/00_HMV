using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    public enum DwgConvertAction
    {
        LinesAndTexts,
        FilledRegions
    }

    public class DwgConvertWindow : Window
    {
        private ListBox listBox;
        private TextBox searchBox;
        private ComboBox fontCombo;
        private List<string> allItems;

        public List<int> SelectedIndices { get; private set; } = new List<int>();
        public string SelectedFont { get; private set; } = "Romans";
        public DwgConvertAction Action { get; private set; }

        public DwgConvertWindow(List<string> items)
        {
            allItems = items;
            Title = "HMV Tools - DWG Convert";
            Width = 560;
            Height = 560;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(Color.FromRgb(245, 245, 248));

            var mainGrid = new Grid();
            mainGrid.Margin = new Thickness(20);
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // HEADER
            var titleText = new TextBlock
            {
                Text = "Select DWG Links to Convert",
                FontWeight = FontWeights.Bold,
                FontSize = 16,
                Margin = new Thickness(0, 0, 0, 15),
                Foreground = new SolidColorBrush(Color.FromRgb(40, 40, 50))
            };
            Grid.SetRow(titleText, 0);
            mainGrid.Children.Add(titleText);

            // SEARCH & LIST
            var listGrid = new Grid();
            Grid.SetRow(listGrid, 1);
            listGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            listGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            searchBox = new TextBox
            {
                Margin = new Thickness(0, 0, 0, 10),
                Padding = new Thickness(8),
                Background = Brushes.White
            };
            searchBox.TextChanged += SearchBox_TextChanged;
            Grid.SetRow(searchBox, 0);
            listGrid.Children.Add(searchBox);

            listBox = new ListBox
            {
                SelectionMode = SelectionMode.Multiple,
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(220, 220, 225)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(5)
            };
            foreach (string item in items) listBox.Items.Add(CreateListItem(item));
            Grid.SetRow(listBox, 1);
            listGrid.Children.Add(listBox);

            mainGrid.Children.Add(listGrid);

            // BOTTOM PANEL (Font & 2 Buttons)
            var bottomPanel = new StackPanel { Orientation = Orientation.Vertical, Margin = new Thickness(0, 20, 0, 0) };
            Grid.SetRow(bottomPanel, 2);

            var fontPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 15) };
            fontPanel.Children.Add(new TextBlock { Text = "Text Font:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 10, 0) });
            fontCombo = new ComboBox { Width = 150, VerticalContentAlignment = VerticalAlignment.Center };
            fontCombo.Items.Add("Romans");
            fontCombo.Items.Add("Arial");
            fontCombo.Items.Add("ISOCPEUR");
            fontCombo.SelectedIndex = 0;
            fontPanel.Children.Add(fontCombo);
            bottomPanel.Children.Add(fontPanel);

            var buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center };

            // BUTTON 1: Lines & Texts (Opus stable logic)
            var btnLinesTexts = new Button
            {
                Content = "Convert Lines & Texts",
                Foreground = Brushes.White,
                FontWeight = FontWeights.Bold,
                FontSize = 14,
                Cursor = Cursors.Hand,
                Template = CreateButtonTemplate(Color.FromRgb(0, 120, 215)),
                Margin = new Thickness(0, 0, 10, 0)
            };
            btnLinesTexts.Click += (s, e) => HandleExecution(DwgConvertAction.LinesAndTexts);
            buttonPanel.Children.Add(btnLinesTexts);

            // BUTTON 2: Filled Regions
            var btnRegions = new Button
            {
                Content = "Convert Filled Regions",
                Foreground = Brushes.White,
                FontWeight = FontWeights.Bold,
                FontSize = 14,
                Cursor = Cursors.Hand,
                Template = CreateButtonTemplate(Color.FromRgb(40, 167, 69))
            };
            btnRegions.Click += (s, e) => HandleExecution(DwgConvertAction.FilledRegions);
            buttonPanel.Children.Add(btnRegions);

            bottomPanel.Children.Add(buttonPanel);
            mainGrid.Children.Add(bottomPanel);
            Content = mainGrid;
        }

        private void HandleExecution(DwgConvertAction action)
        {
            foreach (var item in listBox.SelectedItems)
            {
                string text = ((TextBlock)item).Text;
                SelectedIndices.Add(allItems.IndexOf(text));
            }
            if (SelectedIndices.Count == 0)
            {
                MessageBox.Show("Please select at least one DWG.");
                return;
            }
            SelectedFont = fontCombo.SelectedItem.ToString();
            Action = action;
            DialogResult = true;
        }

        private ControlTemplate CreateButtonTemplate(Color bgColor)
        {
            var template = new ControlTemplate(typeof(Button));
            var border = new FrameworkElementFactory(typeof(Border));
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
            border.SetValue(Border.BackgroundProperty, new SolidColorBrush(bgColor));
            border.SetValue(Border.PaddingProperty, new Thickness(14, 10, 14, 10));

            var content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            content.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);

            border.AppendChild(content);
            template.VisualTree = border;
            return template;
        }

        private TextBlock CreateListItem(string text)
        {
            return new TextBlock { Text = text, Padding = new Thickness(8, 6, 8, 6) };
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            listBox.Items.Clear();
            string filter = searchBox.Text.ToLower();
            foreach (string item in allItems)
            {
                if (item.ToLower().Contains(filter)) listBox.Items.Add(CreateListItem(item));
            }
        }
    }
}