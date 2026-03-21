using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    // ── Data classes (plain, no Revit usings) ──────────────────

    public class TextTypeEntry
    {
        public string TypeName { get; set; }
        public int TypeIdInt { get; set; }
    }

    public class GenericAnnotationTagSettings
    {
        public int TextTypeIdInt { get; set; }
        public double OffsetXMm { get; set; }
        public double OffsetYMm { get; set; }
        public string ParameterName { get; set; }
    }

    // ── Window ─────────────────────────────────────────────────

    public class GenericAnnotationTagWindow : Window
    {
        // Controls
        private TextBox typeSearchBox;
        private ListBox typeListBox;
        private TextBox offsetXBox;
        private TextBox offsetYBox;
        private TextBox paramSearchBox;
        private ListBox paramListBox;

        // Data
        private List<TextTypeEntry> allTypes;
        private List<string> allParameters;

        // Colors
        private static readonly Color BluePrimary = Color.FromRgb(0, 120, 212);
        private static readonly Color GrayBg = Color.FromRgb(240, 240, 243);
        private static readonly Color DarkText = Color.FromRgb(30, 30, 30);
        private static readonly Color MutedText = Color.FromRgb(120, 120, 130);
        private static readonly Color BorderColor = Color.FromRgb(200, 200, 210);
        private static readonly Color WindowBg = Color.FromRgb(245, 245, 248);

        public GenericAnnotationTagSettings Settings { get; private set; }

        public GenericAnnotationTagWindow(
            List<TextTypeEntry> textTypes,
            List<string> parameters)
        {
            allTypes = textTypes;
            allParameters = parameters;

            Title = "HMV Tools – Generic Annotation Tag";
            Width = 520;
            Height = 660;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(WindowBg);

            var main = new Grid();
            main.Margin = new Thickness(20);
            //  0 Title
            //  1 TextType subtitle
            //  2 TextType search
            //  3 TextType list      (Star)
            //  4 Offset panel
            //  5 Param subtitle
            //  6 Param search
            //  7 Param list         (Star)
            //  8 Buttons
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            main.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            main.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // ── Row 0: Title ───────────────────────────────────
            var title = new TextBlock
            {
                Text = "Generic Annotation Tag",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 10)
            };
            Grid.SetRow(title, 0);
            main.Children.Add(title);

            // ── Row 1: TextType subtitle ───────────────────────
            var typeSubtitle = new TextBlock
            {
                Text = "Text note type",
                FontSize = 12,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(0, 0, 0, 4)
            };
            Grid.SetRow(typeSubtitle, 1);
            main.Children.Add(typeSubtitle);

            // ── Row 2: TextType search ─────────────────────────
            typeSearchBox = CreateSearchBox(out Border typeSearchBorder);
            typeSearchBox.TextChanged += TypeSearch_Changed;
            Grid.SetRow(typeSearchBorder, 2);
            main.Children.Add(typeSearchBorder);

            // ── Row 3: TextType list ───────────────────────────
            typeListBox = CreateListBox(out Border typeListBorder);
            typeListBox.MouseDoubleClick += (s, e) =>
            {
                if (typeListBox.SelectedItem != null)
                    paramSearchBox.Focus();
            };
            Grid.SetRow(typeListBorder, 3);
            main.Children.Add(typeListBorder);

            // ── Row 4: Offset panel ────────────────────────────
            var offsetPanel = CreateOffsetPanel();
            Grid.SetRow(offsetPanel, 4);
            main.Children.Add(offsetPanel);

            // ── Row 5: Parameter subtitle ──────────────────────
            var paramSubtitle = new TextBlock
            {
                Text = "Source parameter (from linked element)",
                FontSize = 12,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(0, 0, 0, 4)
            };
            Grid.SetRow(paramSubtitle, 5);
            main.Children.Add(paramSubtitle);

            // ── Row 6: Parameter search ────────────────────────
            paramSearchBox = CreateSearchBox(out Border paramSearchBorder);
            paramSearchBox.TextChanged += ParamSearch_Changed;
            Grid.SetRow(paramSearchBorder, 6);
            main.Children.Add(paramSearchBorder);

            // ── Row 7: Parameter list ──────────────────────────
            paramListBox = CreateListBox(out Border paramListBorder);
            paramListBox.MouseDoubleClick += (s, e) =>
            {
                if (paramListBox.SelectedItem != null) Accept();
            };
            Grid.SetRow(paramListBorder, 7);
            main.Children.Add(paramListBorder);

            // ── Row 8: Buttons ─────────────────────────────────
            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 10, 0, 0)
            };

            var cancelBtn = CreateButton("Cancel", GrayBg,
                Color.FromRgb(60, 60, 60));
            cancelBtn.Width = 90;
            cancelBtn.Margin = new Thickness(0, 0, 8, 0);
            cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };

            var genBtn = CreateButton("Generate", BluePrimary,
                Color.FromRgb(255, 255, 255));
            genBtn.Width = 130;
            genBtn.Click += (s, e) => Accept();

            btnPanel.Children.Add(cancelBtn);
            btnPanel.Children.Add(genBtn);
            Grid.SetRow(btnPanel, 8);
            main.Children.Add(btnPanel);

            Content = main;

            PopulateTypes(allTypes);
            PopulateParameters(allParameters);

            Loaded += (s, e) => typeSearchBox.Focus();
        }

        // ── Accept ─────────────────────────────────────────────

        private void Accept()
        {
            if (typeListBox.SelectedItem == null)
            {
                MessageBox.Show("Select a text note type.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }
            if (paramListBox.SelectedItem == null)
            {
                MessageBox.Show("Select a source parameter.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            double offX = 0, offY = 0;
            if (!string.IsNullOrWhiteSpace(offsetXBox.Text))
                double.TryParse(offsetXBox.Text, out offX);
            if (!string.IsNullOrWhiteSpace(offsetYBox.Text))
                double.TryParse(offsetYBox.Text, out offY);

            string selType =
                (typeListBox.SelectedItem as TextBlock)?.Text ?? "";
            var entry = allTypes.FirstOrDefault(
                t => t.TypeName == selType);

            if (entry == null)
            {
                MessageBox.Show("Could not resolve selected type.",
                    "HMV Tools", MessageBoxButton.OK);
                return;
            }

            string selParam =
                (paramListBox.SelectedItem as TextBlock)?.Text ?? "";

            Settings = new GenericAnnotationTagSettings
            {
                TextTypeIdInt = entry.TypeIdInt,
                OffsetXMm = offX,
                OffsetYMm = offY,
                ParameterName = selParam
            };

            DialogResult = true;
            Close();
        }

        // ── Offset panel ───────────────────────────────────────

        private StackPanel CreateOffsetPanel()
        {
            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 8, 0, 8)
            };

            panel.Children.Add(new TextBlock
            {
                Text = "Offset X (mm):",
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 6, 0)
            });
            offsetXBox = CreateNumericBox("0", out Border bdrX);
            bdrX.Width = 80;
            bdrX.Margin = new Thickness(0, 0, 16, 0);
            panel.Children.Add(bdrX);

            panel.Children.Add(new TextBlock
            {
                Text = "Offset Y (mm):",
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 6, 0)
            });
            offsetYBox = CreateNumericBox("0", out Border bdrY);
            bdrY.Width = 80;
            panel.Children.Add(bdrY);

            return panel;
        }

        // ── Search handlers ────────────────────────────────────

        private void TypeSearch_Changed(object sender,
            TextChangedEventArgs e)
        {
            string filter = typeSearchBox.Text.ToLower();
            typeListBox.Items.Clear();
            foreach (var t in allTypes)
            {
                if (t.TypeName.ToLower().Contains(filter))
                    typeListBox.Items.Add(CreateListItem(t.TypeName));
            }
            if (typeListBox.Items.Count > 0)
                typeListBox.SelectedIndex = 0;
        }

        private void ParamSearch_Changed(object sender,
            TextChangedEventArgs e)
        {
            string filter = paramSearchBox.Text.ToLower();
            paramListBox.Items.Clear();
            foreach (string p in allParameters)
            {
                if (p.ToLower().Contains(filter))
                    paramListBox.Items.Add(CreateListItem(p));
            }
            if (paramListBox.Items.Count > 0)
                paramListBox.SelectedIndex = 0;
        }

        // ── Populate ───────────────────────────────────────────

        private void PopulateTypes(List<TextTypeEntry> entries)
        {
            typeListBox.Items.Clear();
            foreach (var t in entries)
                typeListBox.Items.Add(CreateListItem(t.TypeName));
            if (typeListBox.Items.Count > 0)
                typeListBox.SelectedIndex = 0;
        }

        private void PopulateParameters(List<string> names)
        {
            paramListBox.Items.Clear();
            foreach (string n in names)
                paramListBox.Items.Add(CreateListItem(n));
            if (paramListBox.Items.Count > 0)
                paramListBox.SelectedIndex = 0;
        }

        // ── UI factory helpers ─────────────────────────────────

        private TextBox CreateSearchBox(out Border border)
        {
            border = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 6)
            };
            var box = new TextBox
            {
                Height = 30,
                FontSize = 13,
                VerticalContentAlignment = VerticalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(10, 0, 10, 0)
            };
            Border b = border;
            box.GotFocus += (s, e) =>
                b.BorderBrush = new SolidColorBrush(BluePrimary);
            box.LostFocus += (s, e) =>
                b.BorderBrush = new SolidColorBrush(BorderColor);
            border.Child = box;
            return box;
        }

        private ListBox CreateListBox(out Border border)
        {
            border = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 6)
            };
            var lb = new ListBox
            {
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                FontSize = 13,
                Padding = new Thickness(4),
                SelectionMode = SelectionMode.Single
            };
            border.Child = lb;
            return lb;
        }

        private TextBox CreateNumericBox(string defaultVal,
            out Border border)
        {
            border = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White
            };
            var box = new TextBox
            {
                Text = defaultVal,
                Height = 30,
                FontSize = 13,
                VerticalContentAlignment = VerticalAlignment.Center,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(6, 0, 6, 0)
            };
            Border b = border;
            box.GotFocus += (s, e) =>
                b.BorderBrush = new SolidColorBrush(BluePrimary);
            box.LostFocus += (s, e) =>
                b.BorderBrush = new SolidColorBrush(BorderColor);
            box.PreviewTextInput += (s, e) =>
            {
                foreach (char c in e.Text)
                {
                    if (!char.IsDigit(c) && c != '.' && c != '-')
                    { e.Handled = true; return; }
                    if (c == '.' && box.Text.Contains("."))
                    { e.Handled = true; return; }
                    if (c == '-' && box.SelectionStart > 0)
                    { e.Handled = true; return; }
                    if (c == '-' && box.Text.Contains("-")
                        && !box.SelectedText.Contains("-"))
                    { e.Handled = true; return; }
                }
            };
            border.Child = box;
            return box;
        }

        private TextBlock CreateListItem(string text)
        {
            return new TextBlock
            {
                Text = text,
                Padding = new Thickness(8, 5, 8, 5)
            };
        }

        private Button CreateButton(string text, Color bgColor,
            Color fgColor)
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

        private ControlTemplate GetRoundButtonTemplate(Color bg)
        {
            var tmpl = new ControlTemplate(typeof(Button));
            var border = new FrameworkElementFactory(typeof(Border));
            border.SetValue(Border.CornerRadiusProperty,
                new CornerRadius(6));
            border.SetValue(Border.BackgroundProperty,
                new SolidColorBrush(bg));
            border.SetValue(Border.PaddingProperty,
                new Thickness(14, 6, 14, 6));

            var cp = new FrameworkElementFactory(
                typeof(ContentPresenter));
            cp.SetValue(ContentPresenter.HorizontalAlignmentProperty,
                HorizontalAlignment.Center);
            cp.SetValue(ContentPresenter.VerticalAlignmentProperty,
                VerticalAlignment.Center);

            border.AppendChild(cp);
            tmpl.VisualTree = border;
            return tmpl;
        }
    }
}