using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace HMVTools
{
    public class ConstraintsReleaseWindow : Window
    {
        public ConstraintsReleaseWindow(int selectedCount)
        {
            Title = "HMV Tools - Total Constraints Release";
            Width = 420;
            Height = 300;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(Color.FromRgb(245, 245, 248));

            // --- INTERFAZ GRÁFICA ---
            var mainGrid = new Grid();
            mainGrid.Margin = new Thickness(20);
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // 1. Text Info (Updated for the new logic)
            var txtInfo = new TextBlock
            {
                Text = $"You have {selectedCount} element(s) selected.\n\nThis action will completely unconstrain them by:\n• Unpinning the object.\n• Deleting all Alignment constraints (hidden padlocks).\n• Unlocking all Dimensions, EQs, and Labels (from this object and surrounding items).",
                TextWrapping = TextWrapping.Wrap,
                FontSize = 13,
                Foreground = new SolidColorBrush(Color.FromRgb(50, 50, 50)),
                Margin = new Thickness(0, 0, 0, 20)
            };
            Grid.SetRow(txtInfo, 0);
            mainGrid.Children.Add(txtInfo);

            // 2. Action Button
            Button btnRelease = new Button
            {
                Content = "Release All Constraints",
                Height = 45,
                Foreground = Brushes.White,
                FontSize = 15,
                FontWeight = FontWeights.Bold,
                Template = CreateButtonTemplate(Color.FromRgb(231, 76, 60))
            };
            btnRelease.Click += BtnRelease_Click;
            Grid.SetRow(btnRelease, 2);
            mainGrid.Children.Add(btnRelease);

            this.Content = mainGrid;
        }

        private void BtnRelease_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }

        private ControlTemplate CreateButtonTemplate(Color bgColor)
        {
            var template = new ControlTemplate(typeof(Button));
            var border = new FrameworkElementFactory(typeof(Border));
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
            border.SetValue(Border.BackgroundProperty, new SolidColorBrush(bgColor));

            var content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            content.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);

            border.AppendChild(content);
            template.VisualTree = border;
            return template;
        }
    }
}