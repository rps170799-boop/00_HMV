using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    public class TopographyToLinesWindow : Window
    {
        private ListBox linkListBox;
        private TextBox linkSearchBox;
        private List<string> allLinks;
        private CheckBox chkEtiquetado;

        public int SelectedLinkIndex { get; private set; } = -1;
        public bool GenerateOffsets => chkEtiquetado.IsChecked == true;

        // HMV palette
        private static readonly Color COL_BG = Color.FromRgb(245, 245, 248);
        private static readonly Color COL_ACCENT = Color.FromRgb(0, 120, 212);
        private static readonly Color COL_BORDER = Color.FromRgb(200, 200, 210);
        private static readonly Color COL_TEXT = Color.FromRgb(30, 30, 30);
        private static readonly Color COL_SUB = Color.FromRgb(120, 120, 130);
        private static readonly Color COL_BTN_BG = Color.FromRgb(240, 240, 243);

        public TopographyToLinesWindow(List<string> linkNames)
        {
            allLinks = linkNames;

            Title = "HMV Tools - Topography to Lines (Paso 1)";
            Width = 450;
            Height = 520;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(COL_BG);

            Grid mainGrid = new Grid { Margin = new Thickness(20) };
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 0 Title
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 1 Subtitle
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // 2 List
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 3 Checkbox
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 4 Buttons

            // ── HEADER ──
            TextBlock title = new TextBlock { Text = "1. Vínculo de Topografía", FontSize = 18, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(COL_TEXT), Margin = new Thickness(0, 0, 0, 4) };
            Grid.SetRow(title, 0); mainGrid.Children.Add(title);

            TextBlock subtitle = new TextBlock { Text = "Seleccione el vínculo de Revit que contiene la topografía original.", FontSize = 12, Foreground = new SolidColorBrush(COL_SUB), Margin = new Thickness(0, 0, 0, 16) };
            Grid.SetRow(subtitle, 1); mainGrid.Children.Add(subtitle);

            // ── LINK LIST COLUMN ──
            Grid colGrid = new Grid();
            colGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Search
            colGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // List

            Border searchBorder = new Border { CornerRadius = new CornerRadius(6), BorderBrush = new SolidColorBrush(COL_BORDER), BorderThickness = new Thickness(1), Background = Brushes.White, Margin = new Thickness(0, 0, 0, 8) };
            linkSearchBox = new TextBox { Height = 28, FontSize = 12, BorderThickness = new Thickness(0), Background = Brushes.Transparent, Padding = new Thickness(8, 0, 8, 0), VerticalContentAlignment = VerticalAlignment.Center, Foreground = new SolidColorBrush(COL_SUB), Text = "Buscar..." };

            linkSearchBox.GotFocus += (s, e) => { if (linkSearchBox.Text == "Buscar...") { linkSearchBox.Text = ""; linkSearchBox.Foreground = new SolidColorBrush(COL_TEXT); } searchBorder.BorderBrush = new SolidColorBrush(COL_ACCENT); };
            linkSearchBox.LostFocus += (s, e) => { if (string.IsNullOrWhiteSpace(linkSearchBox.Text)) { linkSearchBox.Text = "Buscar..."; linkSearchBox.Foreground = new SolidColorBrush(COL_SUB); } searchBorder.BorderBrush = new SolidColorBrush(COL_BORDER); };
            linkSearchBox.TextChanged += (s, e) => { string filter = linkSearchBox.Text; if (filter == "Buscar...") filter = ""; FilterLinkList(filter); };

            searchBorder.Child = linkSearchBox;
            Grid.SetRow(searchBorder, 0); colGrid.Children.Add(searchBorder);

            Border listBorder = new Border { CornerRadius = new CornerRadius(6), BorderBrush = new SolidColorBrush(COL_BORDER), BorderThickness = new Thickness(1), Background = Brushes.White };
            linkListBox = new ListBox { BorderThickness = new Thickness(0), Background = Brushes.Transparent, FontSize = 13, Padding = new Thickness(4), SelectionMode = SelectionMode.Single };
            listBorder.Child = linkListBox;
            Grid.SetRow(listBorder, 1); colGrid.Children.Add(listBorder);

            Grid.SetRow(colGrid, 2); mainGrid.Children.Add(colGrid);
            FilterLinkList(""); // Populate initially

            // ── CHECKBOX ──
            chkEtiquetado = new CheckBox
            {
                Content = "HMV Etiquetado de topografía",
                FontSize = 13,
                Foreground = new SolidColorBrush(COL_TEXT),
                Margin = new Thickness(0, 16, 0, 16),
                VerticalContentAlignment = VerticalAlignment.Center
            };
            Grid.SetRow(chkEtiquetado, 3);
            mainGrid.Children.Add(chkEtiquetado);

            // ── BUTTONS ──
            StackPanel buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            Button cancelBtn = CreateButton("Cancelar", COL_BTN_BG, COL_TEXT, 100);
            cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };
            cancelBtn.Margin = new Thickness(0, 0, 8, 0);
            buttonPanel.Children.Add(cancelBtn);

            Button nextBtn = CreateButton("Siguiente", COL_ACCENT, Colors.White, 140);
            nextBtn.Click += NextBtn_Click;
            buttonPanel.Children.Add(nextBtn);

            Grid.SetRow(buttonPanel, 4); mainGrid.Children.Add(buttonPanel);
            Content = mainGrid;
        }

        private void FilterLinkList(string filter)
        {
            string selectedText = linkListBox?.SelectedItem is TextBlock selectedItem ? selectedItem.Text : null;
            linkListBox.Items.Clear();
            var filteredItems = string.IsNullOrWhiteSpace(filter) ? allLinks : allLinks.FindAll(i => i.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);

            TextBlock itemToSelect = null;
            foreach (var item in filteredItems)
            {
                var listItem = CreateListItem(item);
                linkListBox.Items.Add(listItem);
                if (item == selectedText) itemToSelect = listItem;
            }
            if (itemToSelect != null) linkListBox.SelectedItem = itemToSelect;
            else if (linkListBox.Items.Count == 1) linkListBox.SelectedIndex = 0;
        }

        private void NextBtn_Click(object sender, RoutedEventArgs e)
        {
            if (linkListBox.SelectedIndex < 0)
            {
                MessageBox.Show("Seleccione un vínculo de Revit de la lista.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            string selectedName = ((TextBlock)linkListBox.SelectedItem).Text;
            SelectedLinkIndex = allLinks.IndexOf(selectedName);
            DialogResult = true;
            Close();
        }

        private TextBlock CreateListItem(string text) { return new TextBlock { Text = text, Padding = new Thickness(8, 6, 8, 6), FontSize = 13 }; }

        private Button CreateButton(string text, Color bgColor, Color fgColor, double width)
        {
            Button btn = new Button { Content = text, Width = width, Height = 36, FontSize = 13, Foreground = new SolidColorBrush(fgColor), Background = new SolidColorBrush(bgColor), BorderThickness = new Thickness(0), Cursor = Cursors.Hand };
            btn.Template = GetRoundButtonTemplate(bgColor);
            return btn;
        }

        private ControlTemplate GetRoundButtonTemplate(Color bgColor)
        {
            ControlTemplate template = new ControlTemplate(typeof(Button));
            FrameworkElementFactory border = new FrameworkElementFactory(typeof(Border));
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
            border.SetValue(Border.BackgroundProperty, new SolidColorBrush(bgColor));
            border.SetValue(Border.PaddingProperty, new Thickness(14, 6, 14, 6));
            FrameworkElementFactory content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            content.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            border.AppendChild(content);
            template.VisualTree = border;
            return template;
        }
    }
}