using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    /// <summary>
    /// Read-only report window shown after the Text Audit completes.
    /// Follows the same visual language as PipeAnnotationWindow /
    /// DwgConvertWindow.
    /// </summary>
    public class TextAuditReportWindow : Window
    {
        // Colors (shared palette)
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
        private static readonly Color SuccessGreen =
            Color.FromRgb(22, 163, 74);

        private readonly string reportText;

        public TextAuditReportWindow(string report)
        {
            reportText = report;

            Title = "HMV Tools – Text Audit Report";
            Width = 640;
            Height = 620;
            WindowStartupLocation =
                WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.CanResizeWithGrip;
            Background = new SolidColorBrush(WindowBg);

            var mainGrid = new Grid();
            mainGrid.Margin = new Thickness(20);
            mainGrid.RowDefinitions.Add(
                new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(
                new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(
                new RowDefinition
                {
                    Height = new GridLength(
                        1, GridUnitType.Star)
                });
            mainGrid.RowDefinitions.Add(
                new RowDefinition { Height = GridLength.Auto });

            // ── Row 0: Title ───────────────────────────────────
            var title = new TextBlock
            {
                Text = "Text Audit Complete",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 4)
            };
            Grid.SetRow(title, 0);
            mainGrid.Children.Add(title);

            // ── Row 1: Subtitle ────────────────────────────────
            var subtitle = new TextBlock
            {
                Text = "All text styles, types, and tag " +
                       "families have been standardized.",
                FontSize = 12,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(0, 0, 0, 12)
            };
            Grid.SetRow(subtitle, 1);
            mainGrid.Children.Add(subtitle);

            // ── Row 2: Report text box ─────────────────────────
            var textBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 14)
            };

            var textBox = new TextBox
            {
                Text = report,
                IsReadOnly = true,
                AcceptsReturn = true,
                TextWrapping = TextWrapping.NoWrap,
                VerticalScrollBarVisibility =
                    ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility =
                    ScrollBarVisibility.Auto,
                FontFamily = new FontFamily("Consolas"),
                FontSize = 12,
                Foreground = new SolidColorBrush(DarkText),
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(12, 10, 12, 10)
            };

            textBorder.Child = textBox;
            Grid.SetRow(textBorder, 2);
            mainGrid.Children.Add(textBorder);

            // ── Row 3: Buttons ─────────────────────────────────
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment =
                    HorizontalAlignment.Right
            };

            var copyBtn = CreateButton(
                "Copy to Clipboard", GrayBg,
                Color.FromRgb(60, 60, 60));
            copyBtn.Width = 150;
            copyBtn.Margin = new Thickness(0, 0, 8, 0);
            copyBtn.Click += (s, e) =>
            {
                try
                {
                    Clipboard.SetText(reportText);
                    copyBtn.Content = "Copied ✓";
                }
                catch { /* ignore clipboard errors */ }
            };

            var closeBtn = CreateButton(
                "Close", BluePrimary,
                Color.FromRgb(255, 255, 255));
            closeBtn.Width = 100;
            closeBtn.Click += (s, e) => Close();

            buttonPanel.Children.Add(copyBtn);
            buttonPanel.Children.Add(closeBtn);
            Grid.SetRow(buttonPanel, 3);
            mainGrid.Children.Add(buttonPanel);

            Content = mainGrid;
        }

        // ── UI helpers (same as other HMV windows) ─────────────

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
                ContentPresenter.HorizontalAlignmentProperty,
                HorizontalAlignment.Center);
            content.SetValue(
                ContentPresenter.VerticalAlignmentProperty,
                VerticalAlignment.Center);

            border.AppendChild(content);
            template.VisualTree = border;
            return template;
        }
    }
}