using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    public class SignageSettingsWindow : Window
    {
        public string SelectedSignageType => _cmbSignage.Text;
        public string SelectedNestedFamily => _cmbNested.Text;
        public string SelectedSourceParam => _cmbSource.Text;
        public string SelectedTargetParam => _cmbTarget.Text;

        private ComboBox _cmbSignage;
        private ComboBox _cmbNested;
        private ComboBox _cmbSource;
        private ComboBox _cmbTarget;

        public SignageSettingsWindow(List<string> signageTypes, List<string> nestedNames, List<string> sourceParams, List<string> targetParams)
        {
            Title = "HMV Tools - Signage Hosting";
            Width = 450;
            Height = 480; // Taller to fit the new field
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Background = new SolidColorBrush(Color.FromRgb(245, 245, 248));
            FontFamily = new FontFamily("Segoe UI");

            BuildUI(signageTypes, nestedNames, sourceParams, targetParams);
        }

        private void BuildUI(List<string> signageTypes, List<string> nestedNames, List<string> sourceParams, List<string> targetParams)
        {
            StackPanel root = new StackPanel { Margin = new Thickness(25) };
            Content = root;

            root.Children.Add(new TextBlock
            {
                Text = "Host Signage on Pedestals",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(30, 30, 30)),
                Margin = new Thickness(0, 0, 0, 20)
            });

            // 1. Signage Family
            root.Children.Add(CreateLabel("Select Signage Family Type:"));
            _cmbSignage = CreateSearchableComboBox(signageTypes);
            root.Children.Add(_cmbSignage);

            // 2. Nested Pedestal Family
            root.Children.Add(CreateLabel("Nested Pedestal Family Name:"));
            _cmbNested = CreateSearchableComboBox(nestedNames);
            root.Children.Add(_cmbNested);

            // 3. Source Parameter
            root.Children.Add(CreateLabel("Source Param (Electrical Eq.):"));
            _cmbSource = CreateSearchableComboBox(sourceParams);
            root.Children.Add(_cmbSource);

            // 4. Target Parameter
            root.Children.Add(CreateLabel("Target Param (Signage):"));
            _cmbTarget = CreateSearchableComboBox(targetParams);
            root.Children.Add(_cmbTarget);

            // Buttons
            StackPanel buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 15, 0, 0)
            };

            Button btnCancel = CreateRoundedButton("Cancel", Color.FromRgb(220, 220, 225), Colors.Black, 90);
            btnCancel.Click += (s, e) => { DialogResult = false; Close(); };

            Button btnRun = CreateRoundedButton("Place Signage", Color.FromRgb(0, 120, 212), Colors.White, 120);
            btnRun.Click += (s, e) => { DialogResult = true; Close(); };

            buttonPanel.Children.Add(btnCancel);
            buttonPanel.Children.Add(btnRun);
            root.Children.Add(buttonPanel);
        }

        private ComboBox CreateSearchableComboBox(List<string> items)
        {
            ComboBox cb = new ComboBox
            {
                IsEditable = true,
                StaysOpenOnEdit = true,
                ItemsSource = items,
                Margin = new Thickness(0, 0, 0, 15),
                Height = 30,
                FontSize = 13,
                VerticalContentAlignment = VerticalAlignment.Center
            };

            ICollectionView cv = CollectionViewSource.GetDefaultView(items);
            cb.KeyUp += (s, e) =>
            {
                if (e.Key == Key.Up || e.Key == Key.Down || e.Key == Key.Enter || e.Key == Key.Escape) return;
                string searchText = cb.Text;
                if (string.IsNullOrWhiteSpace(searchText)) cv.Filter = null;
                else cv.Filter = item => (item as string)?.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0;
                cb.IsDropDownOpen = true;
            };

            if (items.Count > 0) cb.SelectedIndex = 0;
            return cb;
        }

        private TextBlock CreateLabel(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = 12,
                FontWeight = FontWeights.Medium,
                Foreground = new SolidColorBrush(Color.FromRgb(90, 90, 100)),
                Margin = new Thickness(0, 0, 0, 4)
            };
        }

        private Button CreateRoundedButton(string text, Color bg, Color fg, double width)
        {
            Button b = new Button
            {
                Content = text,
                Width = width,
                Height = 34,
                FontSize = 13,
                FontWeight = FontWeights.Medium,
                Margin = new Thickness(10, 0, 0, 0),
                Cursor = Cursors.Hand,
                Foreground = new SolidColorBrush(fg),
                BorderThickness = new Thickness(0)
            };

            ControlTemplate template = new ControlTemplate(typeof(Button));
            FrameworkElementFactory border = new FrameworkElementFactory(typeof(Border));
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(5));
            border.SetValue(Border.BackgroundProperty, new SolidColorBrush(bg));

            FrameworkElementFactory content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            content.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);

            border.AppendChild(content);
            template.VisualTree = border;
            b.Template = template;
            return b;
        }
    }
}