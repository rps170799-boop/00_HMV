using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    /// <summary>
    /// Plain data class passed from Command to Window.
    /// No Revit usings needed in this file.
    /// </summary>
    public class SectionViewItem
    {
        public string Name { get; set; }
        public int OriginalIndex { get; set; }
        public bool IsActiveView { get; set; }

        public string DisplayName
        {
            get { return IsActiveView ? Name + "  (Vista Activa)" : Name; }
        }
    }

    public class TopographySectionWindow : Window
    {
        private ListBox viewListBox;
        private TextBox viewSearchBox;
        private List<SectionViewItem> allViews;

        /// <summary>
        /// Indices into the original list passed to the constructor.
        /// </summary>
        public List<int> SelectedViewIndices { get; private set; } = new List<int>();

        private static readonly Color COL_BG = Color.FromRgb(245, 245, 248);
        private static readonly Color COL_ACCENT = Color.FromRgb(0, 120, 212);
        private static readonly Color COL_BORDER = Color.FromRgb(200, 200, 210);
        private static readonly Color COL_TEXT = Color.FromRgb(30, 30, 30);
        private static readonly Color COL_SUB = Color.FromRgb(120, 120, 130);
        private static readonly Color COL_BTN_BG = Color.FromRgb(240, 240, 243);

        public TopographySectionWindow(List<SectionViewItem> sectionViews)
        {
            allViews = sectionViews;

            Title = "HMV Tools - Topography to Lines (Paso 3)";
            Width = 500;
            Height = 560;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(COL_BG);

            Grid mainGrid = new Grid { Margin = new Thickness(20) };
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 0 Title
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 1 Subtitle
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 2 Search
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // 3 List
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 4 Info
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 5 Buttons

            // ── HEADER ──
            TextBlock title = new TextBlock
            {
                Text = "3. Vistas de Sección",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(COL_TEXT),
                Margin = new Thickness(0, 0, 0, 4)
            };
            Grid.SetRow(title, 0);
            mainGrid.Children.Add(title);

            TextBlock subtitle = new TextBlock
            {
                Text = "Seleccione las vistas donde aplicar el script.\nUse Shift o Ctrl para selección múltiple.",
                FontSize = 12,
                Foreground = new SolidColorBrush(COL_SUB),
                Margin = new Thickness(0, 0, 0, 16),
                TextWrapping = TextWrapping.Wrap
            };
            Grid.SetRow(subtitle, 1);
            mainGrid.Children.Add(subtitle);

            // ── SEARCH BAR ──
            Border searchBorder = new Border
            {
                CornerRadius = new CornerRadius(6),
                BorderBrush = new SolidColorBrush(COL_BORDER),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 8)
            };

            viewSearchBox = new TextBox
            {
                Height = 28,
                FontSize = 12,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(8, 0, 8, 0),
                VerticalContentAlignment = VerticalAlignment.Center,
                Foreground = new SolidColorBrush(COL_SUB),
                Text = "Buscar..."
            };

            viewSearchBox.GotFocus += (s, e) =>
            {
                if (viewSearchBox.Text == "Buscar...")
                {
                    viewSearchBox.Text = "";
                    viewSearchBox.Foreground = new SolidColorBrush(COL_TEXT);
                }
                searchBorder.BorderBrush = new SolidColorBrush(COL_ACCENT);
            };

            viewSearchBox.LostFocus += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(viewSearchBox.Text))
                {
                    viewSearchBox.Text = "Buscar...";
                    viewSearchBox.Foreground = new SolidColorBrush(COL_SUB);
                }
                searchBorder.BorderBrush = new SolidColorBrush(COL_BORDER);
            };

            viewSearchBox.TextChanged += (s, e) =>
            {
                string filter = viewSearchBox.Text;
                if (filter == "Buscar...") filter = "";
                FilterViewList(filter);
            };

            searchBorder.Child = viewSearchBox;
            Grid.SetRow(searchBorder, 2);
            mainGrid.Children.Add(searchBorder);

            // ── VIEW LIST (Extended selection = Shift/Ctrl multi-select) ──
            Border listBorder = new Border
            {
                CornerRadius = new CornerRadius(6),
                BorderBrush = new SolidColorBrush(COL_BORDER),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 16)
            };

            viewListBox = new ListBox
            {
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                FontSize = 13,
                Padding = new Thickness(4),
                SelectionMode = SelectionMode.Extended  // Shift + Ctrl multi-select
            };

            listBorder.Child = viewListBox;
            Grid.SetRow(listBorder, 3);
            mainGrid.Children.Add(listBorder);

            // Populate and pre-select active view
            FilterViewList("");
            PreselectActiveView();

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
                Text = "ℹ  Por defecto se aplica solo a la vista activa.\n" +
                       "   Seleccione múltiples secciones para generar un grupo numerado en cada una.",
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

            Button cancelBtn = CreateButton("Atrás", COL_BTN_BG, COL_TEXT, 100);
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

        private void PreselectActiveView()
        {
            for (int i = 0; i < viewListBox.Items.Count; i++)
            {
                if (viewListBox.Items[i] is TextBlock tb)
                {
                    // Find the item tagged as active view
                    foreach (var v in allViews)
                    {
                        if (v.IsActiveView && tb.Text == v.DisplayName)
                        {
                            viewListBox.SelectedItems.Add(tb);
                            return;
                        }
                    }
                }
            }
            // Fallback: select first item
            if (viewListBox.Items.Count > 0)
                viewListBox.SelectedIndex = 0;
        }

        private void FilterViewList(string filter)
        {
            viewListBox.Items.Clear();
            var filtered = string.IsNullOrWhiteSpace(filter)
                ? allViews
                : allViews.FindAll(v => v.DisplayName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);

            foreach (var item in filtered)
            {
                var li = new TextBlock
                {
                    Text = item.DisplayName,
                    Padding = new Thickness(8, 6, 8, 6),
                    FontSize = 13,
                    FontWeight = item.IsActiveView ? FontWeights.SemiBold : FontWeights.Normal
                };
                viewListBox.Items.Add(li);
            }
        }

        private void ConvertBtn_Click(object sender, RoutedEventArgs e)
        {
            if (viewListBox.SelectedItems.Count == 0)
            {
                MessageBox.Show("Seleccione al menos una vista.", "HMV Tools",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SelectedViewIndices.Clear();

            foreach (var sel in viewListBox.SelectedItems)
            {
                if (sel is TextBlock tb)
                {
                    string displayName = tb.Text;
                    foreach (var v in allViews)
                    {
                        if (v.DisplayName == displayName)
                        {
                            SelectedViewIndices.Add(v.OriginalIndex);
                            break;
                        }
                    }
                }
            }

            DialogResult = true;
            Close();
        }

        private Button CreateButton(string text, Color bgColor, Color fgColor, double width)
        {
            Button btn = new Button
            {
                Content = text,
                Width = width,
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