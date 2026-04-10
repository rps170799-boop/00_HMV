using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace HMVTools
{
    public partial class Dwg3DProgressWindow : Window
    {
        public bool IsCancelled { get; private set; }
        private DateTime _startTime;

        public Dwg3DProgressWindow()
        {
            InitializeComponent();
            _startTime = DateTime.Now;
        }

        // ── Window Controls ─────────────────────────────────────────

        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            IsCancelled = true;
            btnCancel.Content = "Cancelling...";
            btnCancel.IsEnabled = false; // Prevent multiple clicks
        }

        // ── Progress Update Methods ─────────────────────────────────

        public void UpdatePhase(string phase) =>
            Pump(() => txtPhase.Text = phase);

        public void UpdateDetail(string detail) =>
            Pump(() => txtDetail.Text = detail);

        public void UpdateStats(string stats) =>
            Pump(() => txtStats.Text = stats);

        public void SetIndeterminate() =>
            Pump(() => progressBar.IsIndeterminate = true);

        public void SetDeterminate(int max) =>
            Pump(() =>
            {
                progressBar.IsIndeterminate = false;
                progressBar.Maximum = max;
                progressBar.Value = 0;
            });

        public void UpdateProgress(int current, int total)
        {
            Pump(() =>
            {
                double pct = total > 0 ? (current * 100.0 / total) : 0;
                progressBar.Value = current; // Because we set Maximum = total

                TimeSpan elapsed = DateTime.Now - _startTime;
                string time = elapsed.TotalSeconds < 60
                    ? $"{elapsed.TotalSeconds:F0}s"
                    : $"{(int)elapsed.TotalMinutes}m {elapsed.Seconds}s";

                txtStats.Text = $"{current} / {total}  ({pct:F0}%)  —  Elapsed: {time}";
            });
        }

        // ── The "Magic" Pump Method ─────────────────────────────────
        /// <summary>
        /// Pumps the WPF dispatcher to keep the window responsive.
        /// This forces the UI to refresh even while Revit is locked doing heavy work.
        /// </summary>
        private void Pump(Action update)
        {
            try
            {
                // Execute the UI change
                Dispatcher.Invoke(update, DispatcherPriority.Render);

                // Force the window to repaint immediately
                Dispatcher.Invoke(new Action(() => { }), DispatcherPriority.ContextIdle);
            }
            catch
            {
                // Silently ignore if the window is closed/closing
            }
        }
    }
}