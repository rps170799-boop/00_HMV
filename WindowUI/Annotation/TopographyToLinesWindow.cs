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

        private ListBox styleListBox;
        private TextBox styleSearchBox;
        private List<string> allStyles;

        // UI Input TextBoxes
        private TextBox txtOffset1;
        private TextBox txtOffset2;
        private TextBox txtOffset3;

        // Public properties to read the values in the Command file
        public int SelectedLinkIndex { get; private set; } = -1;
        public string SelectedLineStyle { get; private set; } // NEW

        public double Offset1 { get; private set; }
        public double Offset2 { get; private set; }
        public double Offset3 { get; private set; }

        // HMV palette
        private static readonly Color COL_BG = Color.FromRgb(245, 245, 248);
        private static readonly Color COL_ACCENT = Color.FromRgb(0, 120, 212);
        private static readonly Color COL_BORDER = Color.FromRgb(200, 200, 210);
        private static readonly Color COL_TEXT = Color.FromRgb(30, 30, 30);
        private static readonly Color COL_SUB = Color.FromRgb(120, 120, 130);
        private static readonly Color COL_BTN_BG = Color.FromRgb(240, 240, 243);

        public TopographyToLinesWindow(List<string> linkNames, List<string> styleNames)
        {
            allLinks = linkNames;
            allStyles = styleNames;

            Title = "HMV Tools - Topography to Lines";
            Width = 900; // Increased width for dual columns
            Height = 560;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(COL_BG);

            Grid mainGrid = new Grid();
            mainGrid.Margin = new Thickness(20);
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 0 Title
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 1 Subtitle
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // 2 Lists
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 3 Inputs
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 4 Info
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 5 Buttons

            // ── HEADER ──
            TextBlock title = new TextBlock
            {
                Text = "Topografía a Líneas",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(COL_TEXT),
                Margin = new Thickness(0, 0, 0, 4)
            };
            Grid.SetRow(title, 0);
            mainGrid.Children.Add(title);

            TextBlock subtitle = new TextBlock
            {
                Text = "Seleccione el vínculo de Revit, los desfases y el estilo para las líneas generadas",
                FontSize = 12,
                Foreground = new SolidColorBrush(COL_SUB),
                Margin = new Thickness(0, 0, 0, 16)
            };
            Grid.SetRow(subtitle, 1);
            mainGrid.Children.Add(subtitle);

            // ── DUAL LIST COLUMNS ──
            Grid listsGrid = new Grid { Margin = new Thickness(0, 0, 0, 16) };
            listsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            listsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) }); // Spacer
            listsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // Left Column: Links
            CreateListColumn(listsGrid, 0, "1. Vínculo de Revit", out linkSearchBox, out linkListBox, allLinks, FilterLinkList);

            // Right Column: Styles
            CreateListColumn(listsGrid, 2, "2. Estilo de Línea (Desfases 2, 3 y 4)", out styleSearchBox, out styleListBox, allStyles, FilterStyleList);

            Grid.SetRow(listsGrid, 2);
            mainGrid.Children.Add(listsGrid);

            // Populate both lists initially
            FilterLinkList("");
            FilterStyleList("");

            // ── OFFSET INPUTS ──
            Grid inputGrid = new Grid { Margin = new Thickness(0, 0, 0, 16) };
            inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            txtOffset1 = CreateInputPanel(inputGrid, "Desfase 2 (mm):", "2100", 0);
            txtOffset2 = CreateInputPanel(inputGrid, "Desfase 3 (mm):", "2300", 1);
            txtOffset3 = CreateInputPanel(inputGrid, "Desfase 4 (mm):", "5000", 2);

            Grid.SetRow(inputGrid, 3);
            mainGrid.Children.Add(inputGrid);

            // ── INFO ──
            Border infoCard = new Border
            {
                CornerRadius = new CornerRadius(8),
                Background = new SolidColorBrush(Color.FromRgb(255, 250, 235)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(255, 213, 79)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(12),
                Margin = new Thickness(0, 0, 0, 16)
            };

            TextBlock infoText = new TextBlock
            {
                Text = "ℹ  La línea base y el primer desfase (100mm) usarán <Thin Lines>. Los tres desfases adicionales usarán el estilo seleccionado.",
                FontSize = 11,
                Foreground = new SolidColorBrush(Color.FromRgb(102, 60, 0)),
                TextWrapping = TextWrapping.Wrap
            };
            infoCard.Child = infoText;
            Grid.SetRow(infoCard, 4);
            mainGrid.Children.Add(infoCard);

            // ── BUTTONS ──
            StackPanel buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            Button cancelBtn = CreateButton("Cancelar", COL_BTN_BG, COL_TEXT, 100);
            cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };
            cancelBtn.Margin = new Thickness(0, 0, 8, 0);
            buttonPanel.Children.Add(cancelBtn);

            Button convertBtn = CreateButton("Convertir", COL_ACCENT, Colors.White, 140);
            convertBtn.Click += ConvertBtn_Click;
            buttonPanel.Children.Add(convertBtn);

            Grid.SetRow(buttonPanel, 5);
            mainGrid.Children.Add(buttonPanel);

            Content = mainGrid;
        }

        private void CreateListColumn(Grid parentGrid, int colIndex, string headerText, out TextBox searchBox, out ListBox listBox, List<string> sourceList, Action<string> filterAction)
        {
            Grid colGrid = new Grid();
            colGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Header
            colGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Search
            colGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // List

            TextBlock header = new TextBlock { Text = headerText, FontSize = 12, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(COL_TEXT), Margin = new Thickness(0, 0, 0, 8) };
            Grid.SetRow(header, 0);
            colGrid.Children.Add(header);

            Border searchBorder = new Border { CornerRadius = new CornerRadius(6), BorderBrush = new SolidColorBrush(COL_BORDER), BorderThickness = new Thickness(1), Background = Brushes.White, Margin = new Thickness(0, 0, 0, 8) };
            TextBox sBox = new TextBox { Height = 28, FontSize = 12, BorderThickness = new Thickness(0), Background = Brushes.Transparent, Padding = new Thickness(8, 0, 8, 0), VerticalContentAlignment = VerticalAlignment.Center, Foreground = new SolidColorBrush(COL_SUB), Text = "Buscar..." };

            sBox.GotFocus += (s, e) => { if (sBox.Text == "Buscar...") { sBox.Text = ""; sBox.Foreground = new SolidColorBrush(COL_TEXT); } searchBorder.BorderBrush = new SolidColorBrush(COL_ACCENT); };
            sBox.LostFocus += (s, e) => { if (string.IsNullOrWhiteSpace(sBox.Text)) { sBox.Text = "Buscar..."; sBox.Foreground = new SolidColorBrush(COL_SUB); } searchBorder.BorderBrush = new SolidColorBrush(COL_BORDER); };
            sBox.TextChanged += (s, e) => { string filter = sBox.Text; if (filter == "Buscar...") filter = ""; filterAction(filter); };

            searchBorder.Child = sBox;
            Grid.SetRow(searchBorder, 1);
            colGrid.Children.Add(searchBorder);
            searchBox = sBox;

            Border listBorder = new Border { CornerRadius = new CornerRadius(6), BorderBrush = new SolidColorBrush(COL_BORDER), BorderThickness = new Thickness(1), Background = Brushes.White };
            ListBox lBox = new ListBox { BorderThickness = new Thickness(0), Background = Brushes.Transparent, FontSize = 13, Padding = new Thickness(4), SelectionMode = SelectionMode.Single };
            listBorder.Child = lBox;
            Grid.SetRow(listBorder, 2);
            colGrid.Children.Add(listBorder);
            listBox = lBox;

            Grid.SetColumn(colGrid, colIndex);
            parentGrid.Children.Add(colGrid);
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

        private void FilterStyleList(string filter)
        {
            string selectedText = styleListBox?.SelectedItem is TextBlock selectedItem ? selectedItem.Text : null;
            styleListBox.Items.Clear();
            var filteredItems = string.IsNullOrWhiteSpace(filter) ? allStyles : allStyles.FindAll(i => i.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);

            TextBlock itemToSelect = null;
            foreach (var item in filteredItems)
            {
                var listItem = CreateListItem(item);
                styleListBox.Items.Add(listItem);
                if (item == selectedText) itemToSelect = listItem;
            }
            if (itemToSelect != null) styleListBox.SelectedItem = itemToSelect;
            else if (styleListBox.Items.Count == 1) styleListBox.SelectedIndex = 0;
        }

        private void ConvertBtn_Click(object sender, RoutedEventArgs e)
        {
            if (linkListBox.SelectedIndex < 0)
            {
                MessageBox.Show("Seleccione un vínculo de Revit de la lista izquierda.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (styleListBox.SelectedIndex < 0)
            {
                MessageBox.Show("Seleccione un Estilo de Línea de la lista derecha.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Parse text inputs to numbers (defaults to 0 if invalid text)
            double.TryParse(txtOffset1.Text, out double val1);
            double.TryParse(txtOffset2.Text, out double val2);
            double.TryParse(txtOffset3.Text, out double val3);

            Offset1 = val1;
            Offset2 = val2;
            Offset3 = val3;

            // Save selections
            string selectedName = ((TextBlock)linkListBox.SelectedItem).Text;
            SelectedLinkIndex = allLinks.IndexOf(selectedName);

            SelectedLineStyle = ((TextBlock)styleListBox.SelectedItem).Text;

            DialogResult = true;
            Close();
        }

        private TextBox CreateInputPanel(Grid parentGrid, string labelText, string defaultVal, int col)
        {
            StackPanel panel = new StackPanel { Margin = new Thickness(0, 0, 12, 0) };
            if (col == 2) panel.Margin = new Thickness(0);

            TextBlock lbl = new TextBlock { Text = labelText, FontSize = 12, Foreground = new SolidColorBrush(COL_SUB), Margin = new Thickness(0, 0, 0, 4) };
            TextBox tb = new TextBox { Text = defaultVal, Padding = new Thickness(4), Height = 28, VerticalContentAlignment = VerticalAlignment.Center };

            panel.Children.Add(lbl);
            panel.Children.Add(tb);
            Grid.SetColumn(panel, col);
            parentGrid.Children.Add(panel);

            return tb;
        }

        private TextBlock CreateListItem(string text)
        {
            return new TextBlock { Text = text, Padding = new Thickness(8, 6, 8, 6), FontSize = 13 };
        }

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