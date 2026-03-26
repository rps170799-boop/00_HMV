using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    public class TopographyOffsetsWindow : Window
    {
        private ListBox styleListBox;
        private TextBox styleSearchBox;
        private List<string> allStyles;

        private ListBox dimListBox; // NEW: Dimension list
        private TextBox dimSearchBox; // NEW: Dimension search
        private List<string> allDimStyles; // NEW

        private TextBox txtOffset1;
        private TextBox txtOffset2;
        private TextBox txtOffset3;

        private ComboBox cmbTextStyle;

        public string SelectedLineStyle { get; private set; }
        public string SelectedTextStyle { get; private set; }
        public string SelectedDimensionStyle { get; private set; } // NEW

        public double Offset1 { get; private set; }
        public double Offset2 { get; private set; }
        public double Offset3 { get; private set; }

        private static readonly Color COL_BG = Color.FromRgb(245, 245, 248);
        private static readonly Color COL_ACCENT = Color.FromRgb(0, 120, 212);
        private static readonly Color COL_BORDER = Color.FromRgb(200, 200, 210);
        private static readonly Color COL_TEXT = Color.FromRgb(30, 30, 30);
        private static readonly Color COL_SUB = Color.FromRgb(120, 120, 130);
        private static readonly Color COL_BTN_BG = Color.FromRgb(240, 240, 243);

        public TopographyOffsetsWindow(List<string> styleNames, List<string> textStyleNames, List<string> dimStyleNames)
        {
            allStyles = styleNames;
            allDimStyles = dimStyleNames; // NEW

            Title = "HMV Tools - Topography to Lines (Paso 2)";
            Width = 650; // Increased width to fit two lists nicely
            Height = 630;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(COL_BG);

            Grid mainGrid = new Grid { Margin = new Thickness(20) };
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 0 Title
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 1 Subtitle
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // 2 Lists
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 3 Inputs
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 4 Text Style
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 5 Info
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 6 Buttons

            // ── HEADER ──
            TextBlock title = new TextBlock { Text = "2. Estilos y Desfases", FontSize = 18, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(COL_TEXT), Margin = new Thickness(0, 0, 0, 4) };
            Grid.SetRow(title, 0); mainGrid.Children.Add(title);

            TextBlock subtitle = new TextBlock { Text = "Seleccione el estilo de línea y dimensión para los desfases adicionales.", FontSize = 12, Foreground = new SolidColorBrush(COL_SUB), Margin = new Thickness(0, 0, 0, 16) };
            Grid.SetRow(subtitle, 1); mainGrid.Children.Add(subtitle);

            // ── LISTS COLUMN (SIDE BY SIDE) ──
            Grid listsGrid = new Grid { Margin = new Thickness(0, 0, 0, 16) };
            listsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Lines
            listsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(16) }); // Margin
            listsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Dims

            listsGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Labels
            listsGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Searches
            listsGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // ListBoxes

            // Line Style Column
            TextBlock lblLineStyle = new TextBlock { Text = "Estilo de Línea:", FontSize = 12, Foreground = new SolidColorBrush(COL_SUB), Margin = new Thickness(0, 0, 0, 4) };
            Grid.SetColumn(lblLineStyle, 0); Grid.SetRow(lblLineStyle, 0); listsGrid.Children.Add(lblLineStyle);

            Border lineSearchBorder = new Border { CornerRadius = new CornerRadius(6), BorderBrush = new SolidColorBrush(COL_BORDER), BorderThickness = new Thickness(1), Background = Brushes.White, Margin = new Thickness(0, 0, 0, 8) };
            styleSearchBox = new TextBox { Height = 28, FontSize = 12, BorderThickness = new Thickness(0), Background = Brushes.Transparent, Padding = new Thickness(8, 0, 8, 0), VerticalContentAlignment = VerticalAlignment.Center, Foreground = new SolidColorBrush(COL_SUB), Text = "Buscar..." };
            styleSearchBox.GotFocus += (s, e) => { if (styleSearchBox.Text == "Buscar...") { styleSearchBox.Text = ""; styleSearchBox.Foreground = new SolidColorBrush(COL_TEXT); } lineSearchBorder.BorderBrush = new SolidColorBrush(COL_ACCENT); };
            styleSearchBox.LostFocus += (s, e) => { if (string.IsNullOrWhiteSpace(styleSearchBox.Text)) { styleSearchBox.Text = "Buscar..."; styleSearchBox.Foreground = new SolidColorBrush(COL_SUB); } lineSearchBorder.BorderBrush = new SolidColorBrush(COL_BORDER); };
            styleSearchBox.TextChanged += (s, e) => { string filter = styleSearchBox.Text; if (filter == "Buscar...") filter = ""; FilterStyleList(filter); };
            lineSearchBorder.Child = styleSearchBox;
            Grid.SetColumn(lineSearchBorder, 0); Grid.SetRow(lineSearchBorder, 1); listsGrid.Children.Add(lineSearchBorder);

            Border lineListBorder = new Border { CornerRadius = new CornerRadius(6), BorderBrush = new SolidColorBrush(COL_BORDER), BorderThickness = new Thickness(1), Background = Brushes.White };
            styleListBox = new ListBox { BorderThickness = new Thickness(0), Background = Brushes.Transparent, FontSize = 13, Padding = new Thickness(4), SelectionMode = SelectionMode.Single };
            lineListBorder.Child = styleListBox;
            Grid.SetColumn(lineListBorder, 0); Grid.SetRow(lineListBorder, 2); listsGrid.Children.Add(lineListBorder);

            // Dimension Style Column
            TextBlock lblDimStyle = new TextBlock { Text = "Estilo de Dimensión:", FontSize = 12, Foreground = new SolidColorBrush(COL_SUB), Margin = new Thickness(0, 0, 0, 4) };
            Grid.SetColumn(lblDimStyle, 2); Grid.SetRow(lblDimStyle, 0); listsGrid.Children.Add(lblDimStyle);

            Border dimSearchBorder = new Border { CornerRadius = new CornerRadius(6), BorderBrush = new SolidColorBrush(COL_BORDER), BorderThickness = new Thickness(1), Background = Brushes.White, Margin = new Thickness(0, 0, 0, 8) };
            dimSearchBox = new TextBox { Height = 28, FontSize = 12, BorderThickness = new Thickness(0), Background = Brushes.Transparent, Padding = new Thickness(8, 0, 8, 0), VerticalContentAlignment = VerticalAlignment.Center, Foreground = new SolidColorBrush(COL_SUB), Text = "Buscar..." };
            dimSearchBox.GotFocus += (s, e) => { if (dimSearchBox.Text == "Buscar...") { dimSearchBox.Text = ""; dimSearchBox.Foreground = new SolidColorBrush(COL_TEXT); } dimSearchBorder.BorderBrush = new SolidColorBrush(COL_ACCENT); };
            dimSearchBox.LostFocus += (s, e) => { if (string.IsNullOrWhiteSpace(dimSearchBox.Text)) { dimSearchBox.Text = "Buscar..."; dimSearchBox.Foreground = new SolidColorBrush(COL_SUB); } dimSearchBorder.BorderBrush = new SolidColorBrush(COL_BORDER); };
            dimSearchBox.TextChanged += (s, e) => { string filter = dimSearchBox.Text; if (filter == "Buscar...") filter = ""; FilterDimList(filter); };
            dimSearchBorder.Child = dimSearchBox;
            Grid.SetColumn(dimSearchBorder, 2); Grid.SetRow(dimSearchBorder, 1); listsGrid.Children.Add(dimSearchBorder);

            Border dimListBorder = new Border { CornerRadius = new CornerRadius(6), BorderBrush = new SolidColorBrush(COL_BORDER), BorderThickness = new Thickness(1), Background = Brushes.White };
            dimListBox = new ListBox { BorderThickness = new Thickness(0), Background = Brushes.Transparent, FontSize = 13, Padding = new Thickness(4), SelectionMode = SelectionMode.Single };
            dimListBorder.Child = dimListBox;
            Grid.SetColumn(dimListBorder, 2); Grid.SetRow(dimListBorder, 2); listsGrid.Children.Add(dimListBorder);

            Grid.SetRow(listsGrid, 2); mainGrid.Children.Add(listsGrid);
            FilterStyleList(""); // Populate initially
            FilterDimList(""); // Populate initially

            // ── OFFSET INPUTS ──
            Grid inputGrid = new Grid { Margin = new Thickness(0, 0, 0, 16) };
            inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            txtOffset1 = CreateInputPanel(inputGrid, "DISTANCIA DE SEGURIDAD (mm):", "2100", 0);
            txtOffset2 = CreateInputPanel(inputGrid, "VALOR BÁSICO (mm):", "2300", 1);
            txtOffset3 = CreateInputPanel(inputGrid, "NIVEL DE CONEXIÓN (mm):", "5000", 2);

            Grid.SetRow(inputGrid, 3); mainGrid.Children.Add(inputGrid);

            // ── TEXT STYLE COMBOBOX ──
            Grid textGrid = new Grid { Margin = new Thickness(0, 0, 0, 16) };
            textGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Auto) });
            textGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            TextBlock lblText = new TextBlock { Text = "Estilo de Texto:", FontSize = 12, Foreground = new SolidColorBrush(COL_SUB), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 12, 0) };
            Grid.SetColumn(lblText, 0); textGrid.Children.Add(lblText);

            cmbTextStyle = new ComboBox { Height = 28, VerticalContentAlignment = VerticalAlignment.Center, FontSize = 12 };
            foreach (var ts in textStyleNames) cmbTextStyle.Items.Add(ts);
            if (cmbTextStyle.Items.Count > 0) cmbTextStyle.SelectedIndex = 0;
            Grid.SetColumn(cmbTextStyle, 1); textGrid.Children.Add(cmbTextStyle);

            Grid.SetRow(textGrid, 4); mainGrid.Children.Add(textGrid);

            // ── INFO ──
            Border infoCard = new Border { CornerRadius = new CornerRadius(8), Background = new SolidColorBrush(Color.FromRgb(255, 250, 235)), BorderBrush = new SolidColorBrush(Color.FromRgb(255, 213, 79)), BorderThickness = new Thickness(1), Padding = new Thickness(12), Margin = new Thickness(0, 0, 0, 16) };
            TextBlock infoText = new TextBlock { Text = "ℹ  La línea base y el primer desfase usarán <Thin Lines>.\nEl texto y las dimensiones se generarán a la izquierda alineados al Crop Box.", FontSize = 11, Foreground = new SolidColorBrush(Color.FromRgb(102, 60, 0)), TextWrapping = TextWrapping.Wrap };
            infoCard.Child = infoText;
            Grid.SetRow(infoCard, 5); mainGrid.Children.Add(infoCard);

            // ── BUTTONS ──
            StackPanel buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            Button cancelBtn = CreateButton("Atrás", COL_BTN_BG, COL_TEXT, 100);
            cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };
            cancelBtn.Margin = new Thickness(0, 0, 8, 0);
            buttonPanel.Children.Add(cancelBtn);

            Button convertBtn = CreateButton("Convertir", COL_ACCENT, Colors.White, 140);
            convertBtn.Click += ConvertBtn_Click;
            buttonPanel.Children.Add(convertBtn);

            Grid.SetRow(buttonPanel, 6); mainGrid.Children.Add(buttonPanel);
            Content = mainGrid;
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

        private void FilterDimList(string filter)
        {
            string selectedText = dimListBox?.SelectedItem is TextBlock selectedItem ? selectedItem.Text : null;
            dimListBox.Items.Clear();
            var filteredItems = string.IsNullOrWhiteSpace(filter) ? allDimStyles : allDimStyles.FindAll(i => i.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);

            TextBlock itemToSelect = null;
            foreach (var item in filteredItems)
            {
                var listItem = CreateListItem(item);
                dimListBox.Items.Add(listItem);
                if (item == selectedText) itemToSelect = listItem;
            }
            if (itemToSelect != null) dimListBox.SelectedItem = itemToSelect;
            else if (dimListBox.Items.Count == 1) dimListBox.SelectedIndex = 0;
        }

        private void ConvertBtn_Click(object sender, RoutedEventArgs e)
        {
            if (styleListBox.SelectedIndex < 0 || dimListBox.SelectedIndex < 0)
            {
                MessageBox.Show("Seleccione un Estilo de Línea y un Estilo de Dimensión de la lista.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            double.TryParse(txtOffset1.Text, out double val1);
            double.TryParse(txtOffset2.Text, out double val2);
            double.TryParse(txtOffset3.Text, out double val3);

            Offset1 = val1;
            Offset2 = val2;
            Offset3 = val3;

            SelectedLineStyle = ((TextBlock)styleListBox.SelectedItem).Text;
            SelectedDimensionStyle = ((TextBlock)dimListBox.SelectedItem).Text;
            SelectedTextStyle = cmbTextStyle.SelectedItem?.ToString();

            DialogResult = true;
            Close();
        }

        private TextBox CreateInputPanel(Grid parentGrid, string labelText, string defaultVal, int col)
        {
            StackPanel panel = new StackPanel { Margin = new Thickness(0, 0, 12, 0) };
            if (col == 2) panel.Margin = new Thickness(0);

            TextBlock lbl = new TextBlock { Text = labelText, FontSize = 12, Foreground = new SolidColorBrush(COL_SUB), Margin = new Thickness(0, 0, 0, 4) };
            TextBox tb = new TextBox { Text = defaultVal, Padding = new Thickness(4), Height = 28, VerticalContentAlignment = VerticalAlignment.Center };
            panel.Children.Add(lbl); panel.Children.Add(tb);
            Grid.SetColumn(panel, col); parentGrid.Children.Add(panel);
            return tb;
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