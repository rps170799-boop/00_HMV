using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
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
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // Title
            var title = new TextBlock
            {
                Text = "DWG Convert & Standardize",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(30, 30, 30)),
                Margin = new Thickness(0, 0, 0, 4)
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
                Margin = new Thickness(0, 8, 0, 8)
            };
            searchBox = new TextBox
            {
                Height = 32,
                FontSize = 13,
                VerticalContentAlignment = VerticalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(10, 0, 10, 0)
            };
            searchBox.TextChanged += SearchBox_TextChanged;
            searchBox.GotFocus += (s, e) =>
                searchBorder.BorderBrush = new SolidColorBrush(
                    Color.FromRgb(0, 120, 212));
            searchBox.LostFocus += (s, e) =>
                searchBorder.BorderBrush = new SolidColorBrush(
                    Color.FromRgb(200, 200, 210));
            searchBorder.Child = searchBox;
            Grid.SetRow(searchBorder, 1);
            mainGrid.Children.Add(searchBorder);

            // DWG List
            var listBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(Color.FromRgb(200, 200, 210)),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 10)
            };
            listBox = new ListBox
            {
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                FontSize = 13,
                Padding = new Thickness(4),
                SelectionMode = SelectionMode.Extended
            };
            listBorder.Child = listBox;
            Grid.SetRow(listBorder, 2);
            mainGrid.Children.Add(listBorder);

            // Font selector row
            var fontPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 12)
            };
            fontPanel.Children.Add(new TextBlock
            {
                Text = "Font for texts:",
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0)
            });

            fontCombo = new ComboBox
            {
                Width = 200,
                Height = 32,
                FontSize = 13,
                SelectedIndex = 0
            };
            fontCombo.Items.Add("Romans");
            fontCombo.Items.Add("Arial");
            fontCombo.Items.Add("Arial Narrow");
            fontCombo.Items.Add("ISOCPEUR");
            fontCombo.Items.Add("Simplex");
            fontCombo.Items.Add("Romanc");
            fontCombo.Items.Add("Romand");
            fontCombo.Items.Add("Txt");
            fontPanel.Children.Add(fontCombo);
            Grid.SetRow(fontPanel, 3);
            mainGrid.Children.Add(fontPanel);

            // Info text
            var infoText = new TextBlock
            {
                Text = "Step 1: Convert Lines  →  reads geometry from the DWG import\n"
                     + "Step 2: Partial Explode the DWG manually (Revit ribbon)\n"
                     + "Step 3: Standardize All  →  re-styles lines + texts and purges\n"
                     + "        DWG-imported styles (font selector applies here)",
                FontSize = 11,
                Foreground = new SolidColorBrush(Color.FromRgb(120, 120, 130)),
                Margin = new Thickness(0, 0, 0, 12),
                TextWrapping = TextWrapping.Wrap
            };
            Grid.SetRow(infoText, 4);
            mainGrid.Children.Add(infoText);

            // Buttons
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var cancelBtn = CreateButton("Cancel",
                Color.FromRgb(240, 240, 243),
                Color.FromRgb(60, 60, 60));
            cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };
            cancelBtn.Margin = new Thickness(0, 0, 8, 0);
            cancelBtn.Width = 90;

            var linesBtn = CreateButton("Convert Lines",
                Color.FromRgb(0, 120, 212),
                Color.FromRgb(255, 255, 255));
            linesBtn.Click += (s, e) => AcceptWith(DwgConvertAction.ConvertLines);
            linesBtn.Margin = new Thickness(0, 0, 8, 0);
            linesBtn.Width = 130;

            var textsBtn = CreateButton("Standardize All",
                Color.FromRgb(40, 167, 69),
                Color.FromRgb(255, 255, 255));
           
            textsBtn.Width = 150;

            buttonPanel.Children.Add(cancelBtn);
            buttonPanel.Children.Add(linesBtn);
            buttonPanel.Children.Add(textsBtn);
            Grid.SetRow(buttonPanel, 5);
            mainGrid.Children.Add(buttonPanel);

            Content = mainGrid;

            foreach (string item in allItems)
                listBox.Items.Add(CreateListItem(item));

            Loaded += (s, e) => searchBox.Focus();
        }

        private void AcceptWith(DwgConvertAction action)
        {
            // Lines requires DWG selection, Texts doesn't
            if (action == DwgConvertAction.ConvertLines)
            {
                if (listBox.SelectedItems.Count == 0)
                {
                    MessageBox.Show("Select at least one DWG.",
                        "HMV Tools", MessageBoxButton.OK);
                    return;
                }
                SelectedIndices.Clear();
                foreach (var item in listBox.SelectedItems)
                {
                    if (item is TextBlock tb)
                        SelectedIndices.Add(allItems.IndexOf(tb.Text));
                }
            }

            SelectedFont = fontCombo.SelectedItem?.ToString() ?? "Romans";
            Action = action;
            DialogResult = true;
            Close();
        }

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
            border.SetValue(Border.BackgroundProperty,
                new SolidColorBrush(bgColor));
            border.SetValue(Border.PaddingProperty,
                new Thickness(14, 6, 14, 6));

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

        private void SearchBox_TextChanged(object sender,
            TextChangedEventArgs e)
        {
            listBox.Items.Clear();
            string filter = searchBox.Text.ToLower();
            foreach (string item in allItems)
            {
                if (item.ToLower().Contains(filter))
                    listBox.Items.Add(CreateListItem(item));
            }
        }
    }
}