using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;

namespace HMVTools
{
    public partial class ReshapeFlexPipeWindow : Window
    {
        private readonly UIApplication          _uiapp;
        private readonly ReshapeFlexPipeHandler _handler;
        private readonly ExternalEvent          _exEvent;

        private readonly ReshapePickHandler _pickHandler;
        private readonly ExternalEvent      _pickEvent;

        private ElementId       _flexPipeId;
        private List<ElementId> _curveIds = new List<ElementId>();

        public ReshapeFlexPipeWindow(
            UIApplication uiapp,
            ReshapeFlexPipeHandler handler,
            ExternalEvent exEvent)
        {
            InitializeComponent();

            _uiapp   = uiapp;
            _handler = handler;
            _exEvent = exEvent;
            _handler.UI = this;

            _pickHandler = new ReshapePickHandler { UI = this };
            _pickEvent   = ExternalEvent.Create(_pickHandler);

            this.Closed += (s, e) => ReshapeFlexPipeCommand.ClearWindow();
        }

        // ── Called by ReshapePickHandler after a FlexPipe is picked ──────────
        public void OnFlexPipePicked(FlexPipe pipe)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (pipe == null) return;
                _flexPipeId = pipe.Id;
                lblFlexPipe.Text       = $"FlexPipe  Id: {pipe.Id}  ({pipe.FlexPipeType?.Name ?? "?"})";
                lblFlexPipe.Foreground = GreenBrush();
                txtStatus.Text         = "FlexPipe selected ✔  — now pick the path curve(s).";
                this.Show();
                this.Activate();
            }));
        }

        // ── Called by ReshapePickHandler after curves are picked ──────────────
        public void OnCurvesPicked(int count)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                _curveIds = new List<ElementId>(_pickHandler.PickedCurveIds);
                lblCurves.Text       = $"{count} curve(s) selected ✔";
                lblCurves.Foreground = GreenBrush();
                txtStatus.Text       = "Curve(s) selected ✔  — click Reshape FlexPipe to proceed.";
                this.Show();
                this.Activate();
            }));
        }

        public void SetStatus(string message)
        {
            Dispatcher.BeginInvoke(new Action(() => txtStatus.Text = message));
        }

        public void RestoreWindow()
        {
            Dispatcher.BeginInvoke(new Action(() => { this.Show(); this.Activate(); }));
        }

        // ── Button handlers ───────────────────────────────────────────────────

        private void BtnPickFlexPipe_Click(object sender, RoutedEventArgs e)
        {
            _pickHandler.Step = ReshapePickHandler.PickStep.FlexPipe;
            this.Hide();
            _pickEvent.Raise();
        }

        private void BtnPickCurves_Click(object sender, RoutedEventArgs e)
        {
            _pickHandler.Step = ReshapePickHandler.PickStep.Curves;
            this.Hide();
            _pickEvent.Raise();
        }

        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            if (_flexPipeId == null)
            { SetStatus("Error: Select a FlexPipe first."); return; }

            if (_curveIds == null || _curveIds.Count == 0)
            { SetStatus("Error: Select at least one path curve."); return; }

            if (!double.TryParse(txtInterval.Text,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out double interval) || interval <= 0)
            { SetStatus("Error: Resample interval must be a positive number."); return; }

            _handler.FlexPipeId       = _flexPipeId;
            _handler.CurveIds         = _curveIds;
            _handler.ResampleInterval = interval;

            SetStatus("Processing…");
            _exEvent.Raise();
        }

        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => DragMove();

        private void BtnClose_Click(object sender, RoutedEventArgs e) => this.Close();

        private static System.Windows.Media.SolidColorBrush GreenBrush() =>
            new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(0, 130, 60));
    }
}