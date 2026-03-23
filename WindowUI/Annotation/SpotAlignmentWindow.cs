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
        /// <summary>True = move leader (text follows). False = move text only.</summary>
        public bool MoveWithLeader { get; set; }
    }

    // ── Window ─────────────────────────────────────────────────

    public class SpotAlignmentWindow : Window
    {
        // Controls
        private CheckBox chkMoveLeader;

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
            Width = 420;
            SizeToContent = SizeToContent.Height;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(WindowBg);

            var main = new Grid { Margin = new Thickness(24) };
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 0 Title
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 1 Selection info
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 2 Pick line + checkbox
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 3 Info
            main.RowDefinitions.Add(new RowDefinition { Height = new GridLength(20) }); // 4 Spacer
            main.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 5 Buttons

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

            // ── Row 2: Pick line instruction + Move with leader ─
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
                Text = "Pick a Reference Line on screen",
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 4)
            });
            refPanel.Children.Add(new TextBlock
            {
                Text = "After clicking OK, select a model or detail line.\n"
                     + "The alignment axis is detected automatically from\n"
                     + "the line's orientation.",
                FontSize = 11,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(0, 0, 0, 10)
            });
            chkMoveLeader = new CheckBox
            {
                Content = "Move with leader",
                FontSize = 13,
                Foreground = new SolidColorBrush(DarkText),
                IsChecked = false
            };
            refPanel.Children.Add(chkMoveLeader);
            refPanel.Children.Add(new TextBlock
            {
                Text = "Unchecked: moves the text only.\n"
                     + "Checked: moves the leader shoulder (text follows).",
                FontSize = 11,
                Foreground = new SolidColorBrush(MutedText),
                Margin = new Thickness(20, 4, 0, 0)
            });
            refBorder.Child = refPanel;
            Grid.SetRow(refBorder, 2);
            main.Children.Add(refBorder);

            // ── Row 3: Info text ───────────────────────────────
            var info = new TextBlock
            {
                Text = "Aligns selected spot elevations to a common\n"
                     + "coordinate derived from the picked reference line.",
                FontSize = 11,
                Foreground = new SolidColorBrush(MutedText),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 0)
            };
            Grid.SetRow(info, 3);
            main.Children.Add(info);

            // ── Row 5: Buttons ─────────────────────────────────
            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            var cancelBtn = CreateButton("Cancel", GrayBg, Color.FromRgb(60, 60, 60));
            cancelBtn.Width = 90;
            cancelBtn.Margin = new Thickness(0, 0, 8, 0);
            cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };

            var okBtn = CreateButton("OK – Pick Line", BluePrimary, Colors.White);
            okBtn.Width = 140;
            okBtn.Click += (s, e) => Accept();

            btnPanel.Children.Add(cancelBtn);
            btnPanel.Children.Add(okBtn);
            Grid.SetRow(btnPanel, 5);
            main.Children.Add(btnPanel);

            Content = main;
        }

        // ── Accept ─────────────────────────────────────────────

        private void Accept()
        {
            Settings = new SpotAlignmentSettings
            {
                MoveWithLeader = chkMoveLeader.IsChecked == true
            };
            DialogResult = true;
            Close();
        }

        // ── Helpers ────────────────────────────────────────────

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
