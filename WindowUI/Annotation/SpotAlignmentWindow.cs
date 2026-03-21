using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    // ── Settings returned to the command ────────────────────────

    public class SpotAlignmentSettings
    {
        /// <summary>True = align along X (vertical stack), False = align along Y (horizontal stack).</summary>
        public bool AlignToX { get; set; }
        /// <summary>True = user will pick a reference line on screen.</summary>
        public bool PickReferenceLine { get; set; }
        /// <summary>Typed coordinate value (feet). Only valid when PickReferenceLine is false.</summary>
        public double TypedValue { get; set; }
    }

    // ── Window ─────────────────────────────────────────────────

    public class SpotAlignmentWindow : Window
    {
        // Controls
        private RadioButton rbAlignX;
        private RadioButton rbAlignY;
        private RadioButton rbPickLine;
        private RadioButton rbTypeValue;
        private TextBox txtValue;
        private Border txtValueBorder;

        // Colors (same palette as other HMV windows)
        private static readonly Color BluePrimary = Color.FromRgb(0, 120, 212);
        private static readonly Color GrayBg = Color.FromRgb(240, 240, 243);
        private static readonly Color DarkText = Color.FromRgb(30, 30, 30);
        private static readonly Color MutedText = Color.FromRgb(120, 120, 130);
        private static readonly Color BorderColor = Color.FromRgb(200, 200, 210);
        private static readonly Color WindowBg = Color.FromRgb(245, 245, 248);
        private static readonly Color AccentBg = Color.FromRgb(232, 243, 255);
        private static readonly Color AccentBorder = Color.FromRgb(0, 120, 212);

        /// <summary>User's settings, or null if cancelled.</summary>
        public SpotAlignmentSettings Settings { get; private set; }

        public SpotAlignmentWindow(int preSelectedCount)
        {
            Title = "HMV Tools – Align Spot Elevations";
            Width = 460;
            SizeToContent = SizeToContent.Height;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(WindowBg);

            var main = new Grid { Margin = new Thickness(24) };
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                        // 0 Title
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                        // 1 Selection info
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                        // 2 Axis group
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                        // 3 Reference group
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                        // 4 Info
            main.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });   // 5 Spacer
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });                        // 6 Buttons

            // ── Row 0: Title ───────────────────────────────────
            var title = new TextBlock
            {
                Text = "Align Spot Elevations",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 4)
            };
            Grid.SetRow(title, 0);
            main.Children.Add(title);

            // ── Row 1: Selection info ──────────────────────────
            var selInfo = new TextBlock
            {
                Text = $"{preSelectedCount} spot elevation(s) selected",
                FontSize = 12,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(0, 0, 0, 16)
            };
            Grid.SetRow(selInfo, 1);
            main.Children.Add(selInfo);

            // ── Row 2: Axis group ──────────────────────────────
            var axisBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Padding = new Thickness(14, 12, 14, 12),
                Margin = new Thickness(0, 0, 0, 12)
            };
            var axisPanel = new StackPanel();
            axisPanel.Children.Add(new TextBlock
            {
                Text = "Alignment Axis",
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 8)
            });
            rbAlignX = new RadioButton
            {
                Content = "Align to X-Axis (Vertical Stack)",
                FontSize = 13,
                Foreground = new SolidColorBrush(DarkText),
                IsChecked = true,
                Margin = new Thickness(0, 0, 0, 6)
            };
            rbAlignY = new RadioButton
            {
                Content = "Align to Y-Axis (Horizontal Stack)",
                FontSize = 13,
                Foreground = new SolidColorBrush(DarkText)
            };
            axisPanel.Children.Add(rbAlignX);
            axisPanel.Children.Add(rbAlignY);
            axisBorder.Child = axisPanel;
            Grid.SetRow(axisBorder, 2);
            main.Children.Add(axisBorder);

            // ── Row 3: Reference group ─────────────────────────
            var refBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(AccentBorder),
                BorderThickness = new Thickness(1),
                Background = new SolidColorBrush(AccentBg),
                Padding = new Thickness(14, 12, 14, 12),
                Margin = new Thickness(0, 0, 0, 12)
            };
            var refPanel = new StackPanel();
            refPanel.Children.Add(new TextBlock
            {
                Text = "Reference Target",
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 8)
            });
            rbPickLine = new RadioButton
            {
                Content = "Pick a Reference Line on screen",
                FontSize = 13,
                GroupName = "Reference",
                Foreground = new SolidColorBrush(DarkText),
                IsChecked = true,
                Margin = new Thickness(0, 0, 0, 6)
            };
            rbTypeValue = new RadioButton
            {
                Content = "Type coordinate value (feet):",
                FontSize = 13,
                GroupName = "Reference",
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 6)
            };

            // Value text box (disabled until "Type coordinate" is checked)
            txtValueBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Width = 160,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(20, 0, 0, 0),
                IsEnabled = false,
                Opacity = 0.5
            };
            txtValue = new TextBox
            {
                Text = "0.0",
                Height = 30,
                FontSize = 13,
                VerticalContentAlignment = VerticalAlignment.Center,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(8, 0, 8, 0)
            };
            txtValue.PreviewTextInput += (s, e) =>
            {
                e.Handled = !IsNumericInput(e.Text, txtValue.Text);
            };
            txtValue.GotFocus += (s, e) =>
                txtValueBorder.BorderBrush = new SolidColorBrush(BluePrimary);
            txtValue.LostFocus += (s, e) =>
                txtValueBorder.BorderBrush = new SolidColorBrush(BorderColor);
            txtValueBorder.Child = txtValue;

            // Toggle textbox enabled state based on radio selection
            rbPickLine.Checked += (s, e) =>
            {
                txtValueBorder.IsEnabled = false;
                txtValueBorder.Opacity = 0.5;
            };
            rbTypeValue.Checked += (s, e) =>
            {
                txtValueBorder.IsEnabled = true;
                txtValueBorder.Opacity = 1.0;
                txtValue.Focus();
            };

            refPanel.Children.Add(rbPickLine);
            refPanel.Children.Add(rbTypeValue);
            refPanel.Children.Add(txtValueBorder);
            refBorder.Child = refPanel;
            Grid.SetRow(refBorder, 3);
            main.Children.Add(refBorder);

            // ── Row 4: Info text ───────────────────────────────
            var info = new TextBlock
            {
                Text = "Aligns the text and leader shoulder of each selected\n"
                     + "spot elevation to a common coordinate line.",
                FontSize = 11,
                Foreground = new SolidColorBrush(MutedText),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 4, 0, 0)
            };
            Grid.SetRow(info, 4);
            main.Children.Add(info);

            // ── Row 6: Buttons ─────────────────────────────────
            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            var cancelBtn = CreateButton("Cancel", GrayBg, Color.FromRgb(60, 60, 60));
            cancelBtn.Width = 90;
            cancelBtn.Margin = new Thickness(0, 0, 8, 0);
            cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };

            var okBtn = CreateButton("OK", BluePrimary, Colors.White);
            okBtn.Width = 120;
            okBtn.Click += (s, e) => Accept();

            btnPanel.Children.Add(cancelBtn);
            btnPanel.Children.Add(okBtn);
            Grid.SetRow(btnPanel, 6);
            main.Children.Add(btnPanel);

            Content = main;
        }

        // ── Accept ─────────────────────────────────────────────

        private void Accept()
        {
            bool pickLine = rbPickLine.IsChecked == true;
            double typedVal = 0;

            if (!pickLine)
            {
                if (!double.TryParse(txtValue.Text, out typedVal))
                {
                    MessageBox.Show("Enter a valid numeric value.",
                        "HMV Tools", MessageBoxButton.OK);
                    return;
                }
            }

            Settings = new SpotAlignmentSettings
            {
                AlignToX = rbAlignX.IsChecked == true,
                PickReferenceLine = pickLine,
                TypedValue = typedVal
            };

            DialogResult = true;
            Close();
        }

        // ── Helpers ────────────────────────────────────────────

        private bool IsNumericInput(string newText, string currentText)
        {
            foreach (char c in newText)
            {
                if (!char.IsDigit(c) && c != '.' && c != '-') return false;
                if (c == '.' && currentText.Contains(".")) return false;
                if (c == '-' && currentText.Length > 0) return false;
            }
            return true;
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
            border.SetValue(Border.BackgroundProperty, new SolidColorBrush(bgColor));
            border.SetValue(Border.PaddingProperty, new Thickness(14, 6, 14, 6));

            var content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            content.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);

            border.AppendChild(content);
            template.VisualTree = border;
            return template;
        }
    }
}
