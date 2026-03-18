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
        private ListBox listBox;
        private List<string> allItems;

        public int SelectedLinkIndex { get; private set; } = -1;

        // HMV palette
        private static readonly Color COL_BG = Color.FromRgb(245, 245, 248);
        private static readonly Color COL_ACCENT = Color.FromRgb(0, 120, 212);
        private static readonly Color COL_BORDER = Color.FromRgb(200, 200, 210);
        private static readonly Color COL_TEXT = Color.FromRgb(30, 30, 30);
        private static readonly Color COL_SUB = Color.FromRgb(120, 120, 130);
        private static readonly Color COL_BTN_BG = Color.FromRgb(240, 240, 243);

        public TopographyToLinesWindow(List<string> linkNames)
        {
            allItems = linkNames;
            Title = "HMV Tools - Topography to Lines";
            Width = 540;
            Height = 500;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(COL_BG);

            Grid mainGrid = new Grid();
            mainGrid.Margin = new Thickness(20);
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

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
                Text = "Seleccione el vínculo de Revit que contiene la topografía",
                FontSize = 12,
                Foreground = new SolidColorBrush(COL_SUB),
                Margin = new Thickness(0, 0, 0, 16),
                TextWrapping = TextWrapping.Wrap
            };
            Grid.SetRow(subtitle, 1);
            mainGrid.Children.Add(subtitle);

            // ── LINK LIST ──
            Border listBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(COL_BORDER),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 12)
            };
            
            listBox = new ListBox
            {
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                FontSize = 13,
                Padding = new Thickness(8),
                SelectionMode = SelectionMode.Single
            };
            
            listBorder.Child = listBox;
            Grid.SetRow(listBorder, 2);
            mainGrid.Children.Add(listBorder);

            // Populate list
            foreach (string item in allItems)
                listBox.Items.Add(CreateListItem(item));

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
                Text = "ℹ  Este comando extrae las curvas de nivel de la topografía\n" +
                       "   en la vista activa y las convierte en líneas de detalle de Revit.\n\n" +
                       "   • Asegúrese de que el vínculo contenga elementos de topografía\n" +
                       "   • La vista activa debe ser una vista de plano",
                FontSize = 11,
                Foreground = new SolidColorBrush(Color.FromRgb(102, 60, 0)),
                TextWrapping = TextWrapping.Wrap
            };
            infoCard.Child = infoText;
            Grid.SetRow(infoCard, 3);
            mainGrid.Children.Add(infoCard);

            // ── BUTTONS ──
            StackPanel buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            Button cancelBtn = CreateButton("Cancelar", 
                COL_BTN_BG, COL_TEXT, 100);
            cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };
            cancelBtn.Margin = new Thickness(0, 0, 8, 0);
            buttonPanel.Children.Add(cancelBtn);

            Button convertBtn = CreateButton("Convertir", 
                COL_ACCENT, Colors.White, 140);
            convertBtn.Click += ConvertBtn_Click;
            buttonPanel.Children.Add(convertBtn);

            Grid.SetRow(buttonPanel, 4);
            mainGrid.Children.Add(buttonPanel);

            Content = mainGrid;

            Loaded += (s, e) => listBox.Focus();
        }

        private void ConvertBtn_Click(object sender, RoutedEventArgs e)
        {
            if (listBox.SelectedIndex < 0)
            {
                MessageBox.Show(
                    "Seleccione un vínculo de Revit de la lista.",
                    "HMV Tools",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            SelectedLinkIndex = listBox.SelectedIndex;
            DialogResult = true;
            Close();
        }

        private TextBlock CreateListItem(string text)
        {
            return new TextBlock
            {
                Text = text,
                Padding = new Thickness(8, 6, 8, 6),
                FontSize = 13
            };
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