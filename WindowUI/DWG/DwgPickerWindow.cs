using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    public class DwgPickerWindow : Window
    {
        private ListBox listBox;
        private TextBox searchBox;
        private List<string> allItems;

        public List<int> SelectedIndices { get; private set; } = new List<int>();

        public DwgPickerWindow(List<string> items)
        {
            allItems = items;
            Title = "HMV Tools";
            Width = 520;
            Height = 450;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(Color.FromRgb(245, 245, 248));

            var mainGrid = new Grid();
            mainGrid.Margin = new Thickness(20);
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // Title
            var title = new TextBlock
            {
                Text = "Select DWG to Convert",
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(30, 30, 30)),
                Margin = new Thickness(0, 0, 0, 12)
            };
            Grid.SetRow(title, 0);
            mainGrid.Children.Add(title);

            // Search box
            var searchBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(Color.FromRgb(200, 200, 210)),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 10)
            };
            searchBox = new TextBox
            {
                Height = 34,
                FontSize = 13,
                VerticalContentAlignment = VerticalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(10, 0, 10, 0)
            };
            searchBox.TextChanged += SearchBox_TextChanged;
            searchBox.GotFocus += (s, e) =>
            {
                searchBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(0, 120, 212));
            };
            searchBox.LostFocus += (s, e) =>
            {
                searchBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(200, 200, 210));
            };
            searchBorder.Child = searchBox;
            Grid.SetRow(searchBorder, 1);
            mainGrid.Children.Add(searchBorder);

            // List
            var listBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(Color.FromRgb(200, 200, 210)),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 12)
            };
            listBox = new ListBox
            {
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                FontSize = 13,
                Padding = new Thickness(4),
                SelectionMode = SelectionMode.Extended
            };
            listBox.MouseDoubleClick += ListBox_DoubleClick;
            listBorder.Child = listBox;
            Grid.SetRow(listBorder, 2);
            mainGrid.Children.Add(listBorder);

            // Buttons
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var cancelBtn = CreateButton("Cancel",
                Color.FromRgb(240, 240, 243), Color.FromRgb(60, 60, 60));
            cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };
            cancelBtn.Margin = new Thickness(0, 0, 8, 0);

            var okBtn = CreateButton("Convert",
                Color.FromRgb(0, 120, 212), Color.FromRgb(255, 255, 255));
            okBtn.Click += (s, e) => Accept();

            buttonPanel.Children.Add(cancelBtn);
            buttonPanel.Children.Add(okBtn);
            Grid.SetRow(buttonPanel, 3);
            mainGrid.Children.Add(buttonPanel);

            Content = mainGrid;

            foreach (string item in allItems)
                listBox.Items.Add(CreateListItem(item));

            Loaded += (s, e) => searchBox.Focus();
        }

        private Button CreateButton(string text, Color bgColor, Color fgColor)
        {
            var btn = new Button
            {
                Content = text,
                Width = 100,
                Height = 34,
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
            border.SetValue(Border.PaddingProperty, new Thickness(12, 6, 12, 6));

            var content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(ContentPresenter.HorizontalAlignmentProperty,
                HorizontalAlignment.Center);
            content.SetValue(ContentPresenter.VerticalAlignmentProperty,
                VerticalAlignment.Center);

            border.AppendChild(content);
            template.VisualTree = border;
            return template;
        }

        private TextBlock CreateListItem(string text)
        {
            return new TextBlock
            {
                Text = text,
                Padding = new Thickness(8, 6, 8, 6)
            };
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            listBox.Items.Clear();
            string filter = searchBox.Text.ToLower();
            foreach (string item in allItems)
            {
                if (item.ToLower().Contains(filter))
                    listBox.Items.Add(CreateListItem(item));
            }
        }

        private void ListBox_DoubleClick(object sender, MouseButtonEventArgs e)
        {
            Accept();
        }

        private void Accept()
        {
            if (listBox.SelectedItems.Count > 0)
            {
                SelectedIndices.Clear();
                foreach (var item in listBox.SelectedItems)
                {
                    if (item is TextBlock tb)
                        SelectedIndices.Add(allItems.IndexOf(tb.Text));
                }
                DialogResult = true;
                Close();
            }
        }
    }
}