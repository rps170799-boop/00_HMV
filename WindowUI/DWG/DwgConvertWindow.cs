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

        // Propiedades públicas que leerá el Comando de Revit
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

            // --- INTERFAZ GRÁFICA ---
            var mainGrid = new Grid();
            mainGrid.Margin = new Thickness(20);
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // 1. Buscador
            searchBox = new TextBox
            {
                Margin = new Thickness(0, 0, 0, 10),
                Padding = new Thickness(5),
                FontSize = 14
            };
            searchBox.TextChanged += SearchBox_TextChanged;
            Grid.SetRow(searchBox, 0);
            mainGrid.Children.Add(searchBox);

            // 2. ListBox
            listBox = new ListBox
            {
                SelectionMode = SelectionMode.Extended, // Permitir selección múltiple de DWGs
                Margin = new Thickness(0, 0, 0, 15),
                BorderBrush = new SolidColorBrush(Color.FromRgb(200, 200, 200)),
                Background = Brushes.White
            };
            Grid.SetRow(listBox, 1);
            mainGrid.Children.Add(listBox);

            foreach (string item in allItems)
            {
                listBox.Items.Add(CreateListItem(item));
            }

            // 3. Opciones de Fuente
            var fontPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 15) };
            fontPanel.Children.Add(new TextBlock { Text = "Text Font:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 10, 0), FontSize = 14 });

            fontCombo = new ComboBox { Width = 150, Padding = new Thickness(4), FontSize = 14 };
            fontCombo.Items.Add(new ComboBoxItem { Content = "Romans", IsSelected = true });
            fontCombo.Items.Add(new ComboBoxItem { Content = "Arial" });
            fontCombo.Items.Add(new ComboBoxItem { Content = "ISOCPEUR" });
            fontCombo.Items.Add(new ComboBoxItem { Content = "Technic" });
            fontPanel.Children.Add(fontCombo);

            Grid.SetRow(fontPanel, 2);
            mainGrid.Children.Add(fontPanel);

            // 4. Botones
            var btnPanel = new Grid();
            btnPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            btnPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Button btnLines = new Button
            {
                Content = "1. Convert Lines",
                Margin = new Thickness(0, 0, 5, 0),
                Foreground = Brushes.White,
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Template = CreateButtonTemplate(Color.FromRgb(52, 152, 219))
            };
            btnLines.Click += BtnConvertLines_Click; // Lógica conectada
            Grid.SetColumn(btnLines, 0);
            btnPanel.Children.Add(btnLines);

            Button btnTexts = new Button
            {
                Content = "2. Convert Texts",
                Margin = new Thickness(5, 0, 0, 0),
                Foreground = Brushes.White,
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Template = CreateButtonTemplate(Color.FromRgb(46, 204, 113))
            };
            btnTexts.Click += BtnStandardizeTexts_Click; // Lógica conectada
            Grid.SetColumn(btnTexts, 1);
            btnPanel.Children.Add(btnTexts);

            Grid.SetRow(btnPanel, 3);
            mainGrid.Children.Add(btnPanel);

            this.Content = mainGrid;
        }

        // --- LÓGICA DE EJECUCIÓN DEL COMANDO ---

        private void BtnConvertLines_Click(object sender, RoutedEventArgs e)
        {
            if (SaveSelections())
            {
                this.Action = DwgConvertAction.ConvertLines;
                this.DialogResult = true; // Cierra la UI y permite que Revit ejecute la acción
            }
        }

        private void BtnStandardizeTexts_Click(object sender, RoutedEventArgs e)
        {
            if (SaveSelections())
            {
                this.Action = DwgConvertAction.StandardizeTexts;
                this.DialogResult = true; // Cierra la UI y permite que Revit ejecute la acción
            }
        }

        private bool SaveSelections()
        {
            SelectedIndices.Clear();

            // Esto es súper importante: Mapea los seleccionados a la lista original.
            // Si el usuario buscó "planta", el índice 0 del listbox ya no es el DWG 0.
            foreach (var item in listBox.SelectedItems)
            {
                if (item is TextBlock tb)
                {
                    int originalIndex = allItems.IndexOf(tb.Text);
                    if (originalIndex >= 0)
                    {
                        SelectedIndices.Add(originalIndex);
                    }
                }
            }

            if (SelectedIndices.Count == 0)
            {
                MessageBox.Show("Please select at least one DWG from the list.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (fontCombo.SelectedItem is ComboBoxItem cbi)
            {
                SelectedFont = cbi.Content.ToString();
            }

            return true;
        }

        // --- MÉTODOS VISUALES Y PLANTILLAS ---

        private ControlTemplate CreateButtonTemplate(Color bgColor)
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
    }
}