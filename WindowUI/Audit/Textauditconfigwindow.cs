using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    /// <summary>
    /// Configuration window shown before the Text Audit runs.
    /// Lets the user choose Font, Width Factor, and Name Template.
    /// Returns values via public properties when DialogResult = true.
    /// </summary>
    public class TextAuditConfigWindow : Window
    {
        // ── Colors (shared HMV palette) ───────────────────────
        private static readonly Color BluePrimary =
            Color.FromRgb(0, 120, 212);
        private static readonly Color GrayBg =
            Color.FromRgb(240, 240, 243);
        private static readonly Color DarkText =
            Color.FromRgb(30, 30, 30);
        private static readonly Color MutedText =
            Color.FromRgb(120, 120, 130);
        private static readonly Color BorderColor =
            Color.FromRgb(200, 200, 210);
        private static readonly Color WindowBg =
            Color.FromRgb(245, 245, 248);
        private static readonly Color InfoBg =
            Color.FromRgb(235, 243, 254);
        private static readonly Color InfoBorder =
            Color.FromRgb(180, 210, 245);

        // ── Output properties ─────────────────────────────────
        public string SelectedFont { get; private set; }
        public double SelectedWidth { get; private set; }
        public string SelectedNameTemplate { get; private set; }

        // ── Controls ──────────────────────────────────────────
        private ComboBox _fontCombo;
        private TextBox _widthBox;
        private TextBox _templateBox;

        public TextAuditConfigWindow()
        {
            Title = "HMV Tools – Text Audit Settings";
            Width = 480;
            Height = 420;
            WindowStartupLocation =
                WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(WindowBg);

            var mainStack = new StackPanel();
            mainStack.Margin = new Thickness(24, 20, 24, 20);

            // ── Title ─────────────────────────────────────────
            mainStack.Children.Add(new TextBlock
            {
                Text = "Text Audit Configuration",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 2)
            });

            mainStack.Children.Add(new TextBlock
            {
                Text = "Set the standard properties before " +
                       "running the audit.",
                FontSize = 12,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(0, 0, 0, 18)
            });

            // ── Font ──────────────────────────────────────────
            mainStack.Children.Add(CreateLabel("Font"));

            _fontCombo = new ComboBox
            {
                Height = 32,
                FontSize = 13,
                Margin = new Thickness(0, 0, 0, 12),
                IsEditable = true
            };
            _fontCombo.Items.Add("Arial");
            _fontCombo.Items.Add("Arial Narrow");
            _fontCombo.Items.Add("Calibri");
            _fontCombo.Items.Add("Tahoma");
            _fontCombo.Items.Add("Verdana");
            _fontCombo.Items.Add("Times New Roman");
            _fontCombo.SelectedIndex = 0;
            mainStack.Children.Add(WrapInBorder(_fontCombo));

            // ── Width Factor ──────────────────────────────────
            mainStack.Children.Add(CreateLabel("Width Factor"));

            _widthBox = new TextBox
            {
                Text = "1.0",
                Height = 32,
                FontSize = 13,
                VerticalContentAlignment =
                    VerticalAlignment.Center,
                Padding = new Thickness(8, 0, 8, 0),
                BorderThickness = new Thickness(0)
            };
            mainStack.Children.Add(WrapInBorder(_widthBox));

            // ── Name Template ─────────────────────────────────
            mainStack.Children.Add(CreateLabel("Name Template"));

            _templateBox = new TextBox
            {
                Text = "HMV_General_{0}mm Arial",
                Height = 32,
                FontSize = 13,
                VerticalContentAlignment =
                    VerticalAlignment.Center,
                Padding = new Thickness(8, 0, 8, 0),
                BorderThickness = new Thickness(0)
            };
            mainStack.Children.Add(WrapInBorder(_templateBox));

            // ── Template hint ─────────────────────────────────
            mainStack.Children.Add(new TextBlock
            {
                Text = "{0} will be replaced by the text " +
                       "size in mm (e.g. 2.5)",
                FontSize = 11,
                FontStyle = FontStyles.Italic,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(2, 2, 0, 14)
            });

            // ── Convention reference note ─────────────────────
            var notePanel = new Border
            {
                CornerRadius = new CornerRadius(6),
                Background = new SolidColorBrush(InfoBg),
                BorderBrush = new SolidColorBrush(InfoBorder),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(12, 8, 12, 8),
                Margin = new Thickness(0, 0, 0, 18)
            };

            var noteStack = new StackPanel();

            noteStack.Children.Add(new TextBlock
            {
                Text = "HMV Conventions",
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 4)
            });
            noteStack.Children.Add(new TextBlock
            {
                Text = "CIV :  Width: 1.0   |  Font: Arial\n" +
                       "ELM :  Width: 0.80  |  Font: Arial",
                FontSize = 11,
                FontFamily = new FontFamily("Consolas"),
                Foreground = new SolidColorBrush(
                    Color.FromRgb(60, 60, 70))
            });

            notePanel.Child = noteStack;
            mainStack.Children.Add(notePanel);

            // ── Buttons ───────────────────────────────────────
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment =
                    HorizontalAlignment.Right
            };

            var cancelBtn = CreateButton(
                "Cancel", GrayBg,
                Color.FromRgb(60, 60, 60));
            cancelBtn.Width = 100;
            cancelBtn.Margin = new Thickness(0, 0, 8, 0);
            cancelBtn.Click += (s, e) =>
            {
                DialogResult = false;
                Close();
            };

            var applyBtn = CreateButton(
                "Apply", BluePrimary,
                Color.FromRgb(255, 255, 255));
            applyBtn.Width = 120;
            applyBtn.Click += OnApplyClick;

            buttonPanel.Children.Add(cancelBtn);
            buttonPanel.Children.Add(applyBtn);
            mainStack.Children.Add(buttonPanel);

            Content = mainStack;
        }

        // ── Apply handler ─────────────────────────────────────

        private void OnApplyClick(object sender,
            RoutedEventArgs e)
        {
            // Validate font
            string font = _fontCombo.Text?.Trim();
            if (string.IsNullOrEmpty(font))
            {
                MessageBox.Show(
                    "Please enter a font name.",
                    "Validation",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            // Validate width
            double width;
            if (!double.TryParse(
                _widthBox.Text?.Trim()
                    .Replace(',', '.'),
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo
                    .InvariantCulture,
                out width) ||
                width <= 0 || width > 2.0)
            {
                MessageBox.Show(
                    "Width must be a number between " +
                    "0.01 and 2.0.",
                    "Validation",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            // Validate template
            string template = _templateBox.Text?.Trim();
            if (string.IsNullOrEmpty(template) ||
                !template.Contains("{0}"))
            {
                MessageBox.Show(
                    "The name template must contain {0} " +
                    "as a placeholder for the text size.",
                    "Validation",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            SelectedFont = font;
            SelectedWidth = width;
            SelectedNameTemplate = template;

            DialogResult = true;
            Close();
        }

        // ── UI helpers (shared HMV style) ─────────────────────

        private TextBlock CreateLabel(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = 12,
                FontWeight = FontWeights.Medium,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 4)
            };
        }

        private Border WrapInBorder(UIElement child)
        {
            return new Border
            {
                CornerRadius = new CornerRadius(6),
                BorderBrush =
                    new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 12),
                Child = child
            };
        }

        private Button CreateButton(string text,
            Color bgColor, Color fgColor)
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

        private ControlTemplate GetRoundButtonTemplate(
            Color bgColor)
        {
            var template =
                new ControlTemplate(typeof(Button));
            var border = new FrameworkElementFactory(
                typeof(Border));
            border.SetValue(
                Border.CornerRadiusProperty,
                new CornerRadius(6));
            border.SetValue(
                Border.BackgroundProperty,
                new SolidColorBrush(bgColor));
            border.SetValue(
                Border.PaddingProperty,
                new Thickness(14, 6, 14, 6));

            var content = new FrameworkElementFactory(
                typeof(ContentPresenter));
            content.SetValue(
                ContentPresenter
                    .HorizontalAlignmentProperty,
                HorizontalAlignment.Center);
            content.SetValue(
                ContentPresenter
                    .VerticalAlignmentProperty,
                VerticalAlignment.Center);

            border.AppendChild(content);
            template.VisualTree = border;
            return template;
        }
    }
}