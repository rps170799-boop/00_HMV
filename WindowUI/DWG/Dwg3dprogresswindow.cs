using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

using Color = System.Windows.Media.Color;
using Button = System.Windows.Controls.Button;

namespace HMVTools
{
    public class Dwg3DProgressWindow : Window
    {
        public bool IsCancelled { get; private set; }

        private TextBlock _txtPhase, _txtDetail, _txtStats;
        private ProgressBar _bar;
        private Button _btnCancel;
        private DateTime _startTime;

        static readonly Color CA = Color.FromRgb(0, 120, 212);
        static readonly Color CT = Color.FromRgb(30, 30, 30);
        static readonly Color CS = Color.FromRgb(120, 120, 130);
        static readonly Color CBG = Color.FromRgb(245, 245, 248);

        public Dwg3DProgressWindow()
        {
            Title = "HMV Tools - Converting...";
            Width = 460; Height = 260;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(CBG);
            WindowStyle = WindowStyle.ToolWindow;
            Topmost = true;

            _startTime = DateTime.Now;

            StackPanel root = new StackPanel { Margin = new Thickness(20) };
            Content = root;

            _txtPhase = new TextBlock
            {
                Text = "Initializing...",
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(CT),
                Margin = new Thickness(0, 0, 0, 8),
                TextWrapping = TextWrapping.Wrap
            };
            root.Children.Add(_txtPhase);

            _txtDetail = new TextBlock
            {
                Text = "",
                FontSize = 11,
                Foreground = new SolidColorBrush(CS),
                Margin = new Thickness(0, 0, 0, 12),
                TextWrapping = TextWrapping.Wrap
            };
            root.Children.Add(_txtDetail);

            _bar = new ProgressBar
            {
                Height = 20,
                Minimum = 0,
                Maximum = 100,
                Value = 0,
                Foreground = new SolidColorBrush(CA),
                Margin = new Thickness(0, 0, 0, 8)
            };
            root.Children.Add(_bar);

            _txtStats = new TextBlock
            {
                Text = "",
                FontSize = 10,
                Foreground = new SolidColorBrush(CS),
                Margin = new Thickness(0, 0, 0, 12),
                TextWrapping = TextWrapping.Wrap
            };
            root.Children.Add(_txtStats);

            _btnCancel = new Button
            {
                Content = "Cancel",
                Width = 100,
                Height = 32,
                FontSize = 12,
                Cursor = Cursors.Hand,
                HorizontalAlignment = HorizontalAlignment.Right,
                Background = new SolidColorBrush(Color.FromRgb(220, 53, 69)),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0)
            };
            var tp = new ControlTemplate(typeof(Button));
            var bd = new FrameworkElementFactory(typeof(Border));
            bd.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
            bd.SetValue(Border.BackgroundProperty, _btnCancel.Background);
            bd.SetValue(Border.PaddingProperty, new Thickness(10, 4, 10, 4));
            var cp = new FrameworkElementFactory(typeof(ContentPresenter));
            cp.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            cp.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            bd.AppendChild(cp); tp.VisualTree = bd; _btnCancel.Template = tp;
            _btnCancel.Click += (s, e) =>
            {
                IsCancelled = true;
                _btnCancel.IsEnabled = false;
                _btnCancel.Content = "Cancelling...";
                _txtPhase.Text = "Cancelling... please wait";
            };
            root.Children.Add(_btnCancel);
        }

        public void UpdatePhase(string phase) =>
            Pump(() => _txtPhase.Text = phase);

        public void UpdateDetail(string detail) =>
            Pump(() => _txtDetail.Text = detail);

        public void UpdateProgress(int current, int total)
        {
            Pump(() =>
            {
                double pct = total > 0 ? (current * 100.0 / total) : 0;
                _bar.Value = pct;

                TimeSpan elapsed = DateTime.Now - _startTime;
                string time = elapsed.TotalSeconds < 60
                    ? $"{elapsed.TotalSeconds:F0}s"
                    : $"{(int)elapsed.TotalMinutes}m {elapsed.Seconds}s";

                _txtStats.Text = $"{current} / {total}  ({pct:F0}%)  —  Elapsed: {time}";
            });
        }

        public void UpdateStats(string stats) =>
            Pump(() => _txtStats.Text = stats);

        public void SetIndeterminate() =>
            Pump(() => _bar.IsIndeterminate = true);

        public void SetDeterminate(int max) =>
            Pump(() => { _bar.IsIndeterminate = false; _bar.Maximum = max; _bar.Value = 0; });

        /// <summary>
        /// Pumps the WPF dispatcher to keep the window responsive.
        /// Call this periodically from the main thread.
        /// </summary>
        void Pump(Action update)
        {
            try
            {
                Dispatcher.Invoke(update, DispatcherPriority.Render);
                // Pump pending messages so the window actually repaints
                Dispatcher.Invoke(() => { }, DispatcherPriority.Background);
            }
            catch { }
        }
    }
}