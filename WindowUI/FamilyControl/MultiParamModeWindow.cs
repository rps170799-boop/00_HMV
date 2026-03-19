using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    public class MultiParamModeWindow : Window
    {
        public bool IsCommonMode { get; private set; }

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
        private static readonly Color AccentBg =
            Color.FromRgb(232, 243, 255);

        public MultiParamModeWindow(List<string> familyNames)
        {
            Title = "HMV Tools – Parameter Edit Mode";
            Width = 420;
            Height = 320;
            WindowStartupLocation =
                WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(WindowBg);

            var main = new StackPanel
            {
                Margin = new Thickness(24)
            };
            Content = main;

            main.Children.Add(new TextBlock
            {
                Text = "Multiple families detected",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(DarkText),
                Margin = new Thickness(0, 0, 0, 8)
            });

            main.Children.Add(new TextBlock
            {
                Text = "You selected instances from "
                     + familyNames.Count + " different families.\n"
                     + "Choose how to edit parameters:",
                FontSize = 12,
                Foreground = new SolidColorBrush(MutedText),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 16)
            });

            // Family list
            var listBorder = new Border
            {
                CornerRadius = new CornerRadius(6),
                BorderBrush = new SolidColorBrush(BorderColor),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Padding = new Thickness(10, 8, 10, 8),
                Margin = new Thickness(0, 0, 0, 16)
            };
            var listPanel = new StackPanel();
            foreach (string name in familyNames)
            {
                listPanel.Children.Add(new TextBlock
                {
                    Text = "• " + name,
                    FontSize = 11,
                    Foreground = new SolidColorBrush(DarkText),
                    Margin = new Thickness(0, 1, 0, 1)
                });
            }
            listBorder.Child = listPanel;
            main.Children.Add(listBorder);

            // Buttons
            var btnRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var btnCancel = MakeButton("Cancel",
                GrayBg, Color.FromRgb(60, 60, 60), 90);
            btnCancel.Click += (s, e) =>
            {
                DialogResult = false;
                Close();
            };
            btnRow.Children.Add(btnCancel);

            var btnCommon = MakeButton("Common Parameters",
                AccentBg, BluePrimary, 160);
            btnCommon.Click += (s, e) =>
            {
                IsCommonMode = true;
                DialogResult = true;
                Close();
            };
            btnRow.Children.Add(btnCommon);

            var btnEach = MakeButton("Each Family",
                BluePrimary, Colors.White, 120);
            btnEach.FontWeight = FontWeights.SemiBold;
            btnEach.Click += (s, e) =>
            {
                IsCommonMode = false;
                DialogResult = true;
                Close();
            };
            btnRow.Children.Add(btnEach);

            main.Children.Add(btnRow);
        }

        private Button MakeButton(string text,
            Color bg, Color fg, double w)
        {
            var btn = new Button
            {
                Content = text,
                Width = w,
                Height = 34,
                FontSize = 12,
                Margin = new Thickness(0, 0, 8, 0),
                Cursor = Cursors.Hand,
                Foreground = new SolidColorBrush(fg),
                Background = new SolidColorBrush(bg),
                BorderThickness = new Thickness(0)
            };
            var tp = new ControlTemplate(typeof(Button));
            var bd = new FrameworkElementFactory(typeof(Border));
            bd.SetValue(Border.CornerRadiusProperty,
                new CornerRadius(6));
            bd.SetValue(Border.BackgroundProperty,
                new SolidColorBrush(bg));
            bd.SetValue(Border.PaddingProperty,
                new Thickness(10, 6, 10, 6));
            var cp = new FrameworkElementFactory(
                typeof(ContentPresenter));
            cp.SetValue(
                ContentPresenter.HorizontalAlignmentProperty,
                HorizontalAlignment.Center);
            cp.SetValue(
                ContentPresenter.VerticalAlignmentProperty,
                VerticalAlignment.Center);
            bd.AppendChild(cp);
            tp.VisualTree = bd;
            btn.Template = tp;
            return btn;
        }
    }
}
