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

        private ListBox dimListBox;
        private TextBox dimSearchBox;
        private List<string> allDimStyles;

        private TextBox txtOffset1;
        private TextBox txtOffset2;
        private TextBox txtOffset3;

        private ComboBox cmbTextStyle;
        private ComboBox cmbFilledRegionType;
        private CheckBox chkMaskingRegion;

        public string SelectedLineStyle { get; private set; }
        public string SelectedTextStyle { get; private set; }
        public string SelectedDimensionStyle { get; private set; }
        public string SelectedFilledRegionType { get; private set; }
        public bool GenerateMaskingRegion { get; private set; }

        public double Offset1 { get; private set; }
        public double Offset2 { get; private set; }
        public double Offset3 { get; private set; }

        private static readonly Color COL_BG = Color.FromRgb(245, 245, 248);
        private static readonly Color COL_ACCENT = Color.FromRgb(0, 120, 212);
        private static readonly Color COL_BORDER = Color.FromRgb(200, 200, 210);
        private static readonly Color COL_TEXT = Color.FromRgb(30, 30, 30);
        private static readonly Color COL_SUB = Color.FromRgb(120, 120, 130);
        private static readonly Color COL_BTN_BG = Color.FromRgb(240, 240, 243);

        public TopographyOffsetsWindow(
            List<string> styleNames,
            List<string> textStyleNames,
            List<string> dimStyleNames,
            List<string> filledRegionTypeNames)
        {
            allStyles = styleNames;
            allDimStyles = dimStyleNames;

            Title = "HMV Tools - Topography to Lines (Paso 2)";
            Width = 900;
            Height = 760;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(COL_BG);

            Grid mainGrid = new Grid { Margin = new Thickness(20) };
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 0 Title
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 1 Subtitle
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // 2 Lists
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 3 Inputs
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 4 Text Style
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 5 Filled Region Type
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 6 Masking Region
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 7 Info
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 8 Buttons

            // ── HEADER ──
            TextBlock title = new TextBlock { Text = "2. Estilos y Desfases", FontSize = 18, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(COL_TEXT), Margin = new Thickness(0, 0, 0, 4) };
            Grid.SetRow(title, 0); mainGrid.Children.Add(title);

            TextBlock subtitle = new TextBlock { Text = "Seleccione los estilos de línea y dimensión para los desfases.", FontSize = 12, Foreground = new SolidColorBrush(COL_SUB), Margin = new Thickness(0, 0, 0, 16) };
            Grid.SetRow(subtitle, 1); mainGrid.Children.Add(subtitle);

            // ── TWO COLUMN LISTS ──
            Grid listsContainer = new Grid { Margin = new Thickness(0, 0, 0, 16) };
            listsContainer.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            listsContainer.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) });
            listsContainer.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // LEFT: LINE STYLES
            Grid lineStyleGrid = CreateSearchableList("Estilo de Línea:", allStyles, out styleSearchBox, out styleListBox);
            lineStyleGrid.VerticalAlignment = VerticalAlignment.Stretch;
            Grid.SetColumn(lineStyleGrid, 0);
            listsContainer.Children.Add(lineStyleGrid);

            // RIGHT: DIMENSION STYLES
            Grid dimStyleGrid = CreateSearchableList("Estilo de Dimensión:", allDimStyles, out dimSearchBox, out dimListBox);
            dimStyleGrid.VerticalAlignment = VerticalAlignment.Stretch;
            Grid.SetColumn(dimStyleGrid, 2);
            listsContainer.Children.Add(dimStyleGrid);

            Grid.SetRow(listsContainer, 2); mainGrid.Children.Add(listsContainer);

            // ── WIRE SEARCH EVENTS ──
            styleSearchBox.TextChanged += (s, e) =>
            {
                string filter = styleSearchBox.Text;
                if (filter == "Buscar...") filter = "";
                FilterStyleList(filter);
            };

            dimSearchBox.TextChanged += (s, e) =>
            {
                string filter = dimSearchBox.Text;
                if (filter == "Buscar...") filter = "";
                FilterDimList(filter);
            };

            FilterStyleList("");
            FilterDimList("");

            // ── OFFSET INPUTS ──
            Grid inputGrid = new Grid { Margin = new Thickness(0, 0, 0, 16) };
            inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            txtOffset1 = CreateInputPanel(inputGrid, "DIST. SEGURIDAD (mm, desde 0.10m):", "2000", 0);
            txtOffset2 = CreateInputPanel(inputGrid, "VALOR BÁSICO (mm, desde dist. seg.):", "2300", 1);
            txtOffset3 = CreateInputPanel(inputGrid, "NIVEL CONEXIÓN (mm, desde 0.10m):", "5000", 2);
            Grid.SetRow(inputGrid, 3); mainGrid.Children.Add(inputGrid);

            // ── TEXT STYLE COMBOBOX ──
            Grid textGrid = new Grid { Margin = new Thickness(0, 0, 0, 12) };
            textGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Auto) });
            textGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            TextBlock lblText = new TextBlock { Text = "Estilo de Texto:", FontSize = 12, Foreground = new SolidColorBrush(COL_SUB), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 12, 0) };
            Grid.SetColumn(lblText, 0); textGrid.Children.Add(lblText);

            cmbTextStyle = new ComboBox { Height = 28, VerticalContentAlignment = VerticalAlignment.Center, FontSize = 12 };
            foreach (var ts in textStyleNames) cmbTextStyle.Items.Add(ts);
            if (cmbTextStyle.Items.Count > 0) cmbTextStyle.SelectedIndex = 0;
            Grid.SetColumn(cmbTextStyle, 1); textGrid.Children.Add(cmbTextStyle);
            Grid.SetRow(textGrid, 4); mainGrid.Children.Add(textGrid);

            // ── FILLED REGION TYPE COMBOBOX ──
            Grid filledGrid = new Grid { Margin = new Thickness(0, 0, 0, 12) };
            filledGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Auto) });
            filledGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            TextBlock lblFilled = new TextBlock { Text = "Tipo de Región Rellena:", FontSize = 12, Foreground = new SolidColorBrush(COL_SUB), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 12, 0) };
            Grid.SetColumn(lblFilled, 0); filledGrid.Children.Add(lblFilled);

            cmbFilledRegionType = new ComboBox { Height = 28, VerticalContentAlignment = VerticalAlignment.Center, FontSize = 12 };
            cmbFilledRegionType.Items.Add("(Ninguno)");
            foreach (var frName in filledRegionTypeNames) cmbFilledRegionType.Items.Add(frName);
            cmbFilledRegionType.SelectedIndex = 0;
            Grid.SetColumn(cmbFilledRegionType, 1); filledGrid.Children.Add(cmbFilledRegionType);
            Grid.SetRow(filledGrid, 5); mainGrid.Children.Add(filledGrid);

            // ── MASKING REGION CHECKBOX ──
            chkMaskingRegion = new CheckBox
            {
                Content = "Generar Masking Region (debajo de la topografía)",
                FontSize = 13,
                Foreground = new SolidColorBrush(COL_TEXT),
                Margin = new Thickness(0, 0, 0, 16),
                VerticalContentAlignment = VerticalAlignment.Center
            };
            Grid.SetRow(chkMaskingRegion, 6);
            mainGrid.Children.Add(chkMaskingRegion);

            // ── INFO ──
            Border infoCard = new Border { CornerRadius = new CornerRadius(8), Background = new SolidColorBrush(Color.FromRgb(255, 250, 235)), BorderBrush = new SolidColorBrush(Color.FromRgb(255, 213, 79)), BorderThickness = new Thickness(1), Padding = new Thickness(12), Margin = new Thickness(0, 0, 0, 16) };
            TextBlock infoText = new TextBlock
            {
                Text = "ℹ  Se crearán: línea base + desfase 0.10m (<Thin Lines>) + 3 desfases con el estilo seleccionado.\n" +
                       "   Las dimensiones se crearán entre los desfases. Los textos se posicionarán a la izquierda.\n" +
                       "   La región rellena se genera entre la línea base y el desfase de 0.10m.\n" +
                       "   La masking region oculta todo lo que quede debajo de la topografía.",
                FontSize = 11,
                Foreground = new SolidColorBrush(Color.FromRgb(102, 60, 0)),
                TextWrapping = TextWrapping.Wrap
            };
            infoCard.Child = infoText;
            Grid.SetRow(infoCard, 7); mainGrid.Children.Add(infoCard);

            // ── BUTTONS ──
            StackPanel buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            Button cancelBtn = CreateButton("Atrás", COL_BTN_BG, COL_TEXT, 100);
            cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };
            cancelBtn.Margin = new Thickness(0, 0, 8, 0);
            buttonPanel.Children.Add(cancelBtn);

            Button convertBtn = CreateButton("Siguiente", COL_ACCENT, Colors.White, 140);
            convertBtn.Click += ConvertBtn_Click;
            buttonPanel.Children.Add(convertBtn);

            Grid.SetRow(buttonPanel, 8); mainGrid.Children.Add(buttonPanel);
            Content = mainGrid;
        }

        private Grid CreateSearchableList(string label, List<string> source, out TextBox searchBox, out ListBox listBox)
        {
            Grid container = new Grid();
            container.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            container.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            container.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            TextBlock lbl = new TextBlock { Text = label, FontSize = 12, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(COL_SUB), Margin = new Thickness(0, 0, 0, 4) };
            Grid.SetRow(lbl, 0); container.Children.Add(lbl);

            Border searchBorder = new Border { CornerRadius = new CornerRadius(6), BorderBrush = new SolidColorBrush(COL_BORDER), BorderThickness = new Thickness(1), Background = Brushes.White, Margin = new Thickness(0, 0, 0, 8) };
            searchBox = new TextBox { Height = 28, FontSize = 12, BorderThickness = new Thickness(0), Background = Brushes.Transparent, Padding = new Thickness(8, 0, 8, 0), VerticalContentAlignment = VerticalAlignment.Center, Foreground = new SolidColorBrush(COL_SUB), Text = "Buscar..." };

            TextBox innerSearch = searchBox;
            Border innerBorder = searchBorder;
            searchBox.GotFocus += (s, e) => { if (innerSearch.Text == "Buscar...") { innerSearch.Text = ""; innerSearch.Foreground = new SolidColorBrush(COL_TEXT); } innerBorder.BorderBrush = new SolidColorBrush(COL_ACCENT); };
            searchBox.LostFocus += (s, e) => { if (string.IsNullOrWhiteSpace(innerSearch.Text)) { innerSearch.Text = "Buscar..."; innerSearch.Foreground = new SolidColorBrush(COL_SUB); } innerBorder.BorderBrush = new SolidColorBrush(COL_BORDER); };

            searchBorder.Child = searchBox;
            Grid.SetRow(searchBorder, 1); container.Children.Add(searchBorder);

            Border listBorder = new Border { CornerRadius = new CornerRadius(6), BorderBrush = new SolidColorBrush(COL_BORDER), BorderThickness = new Thickness(1), Background = Brushes.White };
            listBox = new ListBox { BorderThickness = new Thickness(0), Background = Brushes.Transparent, FontSize = 13, Padding = new Thickness(4), SelectionMode = SelectionMode.Single };
            listBorder.Child = listBox;
            Grid.SetRow(listBorder, 2); container.Children.Add(listBorder);

            return container;
        }

        private void FilterStyleList(string filter)
        {
            string prev = styleListBox?.SelectedItem is TextBlock tb ? tb.Text : null;
            styleListBox.Items.Clear();
            var filtered = string.IsNullOrWhiteSpace(filter) ? allStyles : allStyles.FindAll(i => i.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);
            TextBlock toSelect = null;
            foreach (var item in filtered)
            {
                var li = CreateListItem(item);
                styleListBox.Items.Add(li);
                if (item == prev) toSelect = li;
            }
            if (toSelect != null) styleListBox.SelectedItem = toSelect;
            else if (styleListBox.Items.Count > 0) styleListBox.SelectedIndex = 0;
        }

        private void FilterDimList(string filter)
        {
            string prev = dimListBox?.SelectedItem is TextBlock tb ? tb.Text : null;
            dimListBox.Items.Clear();
            var filtered = string.IsNullOrWhiteSpace(filter) ? allDimStyles : allDimStyles.FindAll(i => i.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);
            TextBlock toSelect = null;
            foreach (var item in filtered)
            {
                var li = CreateListItem(item);
                dimListBox.Items.Add(li);
                if (item == prev) toSelect = li;
            }
            if (toSelect != null) dimListBox.SelectedItem = toSelect;
            else if (dimListBox.Items.Count > 0) dimListBox.SelectedIndex = 0;
        }

        private void ConvertBtn_Click(object sender, RoutedEventArgs e)
        {
            if (styleListBox.SelectedIndex < 0 || dimListBox.SelectedIndex < 0)
            {
                MessageBox.Show("Seleccione un Estilo de Línea y un Estilo de Dimensión.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            double.TryParse(txtOffset1.Text, out double val1);
            double.TryParse(txtOffset2.Text, out double val2);
            double.TryParse(txtOffset3.Text, out double val3);

            Offset1 = val1; Offset2 = val2; Offset3 = val3;
            SelectedLineStyle = ((TextBlock)styleListBox.SelectedItem).Text;
            SelectedDimensionStyle = ((TextBlock)dimListBox.SelectedItem).Text;
            SelectedTextStyle = cmbTextStyle.SelectedItem?.ToString();

            // Filled region type: null if "(Ninguno)" selected
            string frSelection = cmbFilledRegionType.SelectedItem?.ToString();
            SelectedFilledRegionType = (frSelection == "(Ninguno)") ? null : frSelection;

            GenerateMaskingRegion = chkMaskingRegion.IsChecked == true;

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