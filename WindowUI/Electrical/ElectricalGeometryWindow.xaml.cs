using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using System.Windows.Forms;

namespace HMVTools
{
    public partial class ElectricalGeometryWindow : Window
    {
        private readonly ElectricalGeometryHandler _handler;
        private readonly ExternalEvent             _exEvent;

        public ElectricalGeometryWindow(
            ElectricalGeometryHandler handler,
            ExternalEvent             exEvent,
            List<string>              parameterNames)
        {
            InitializeComponent();

            _handler = handler;
            _exEvent = exEvent;
            _handler.UI = this;

            cmbParameter.ItemsSource = parameterNames;

            int idx = parameterNames.IndexOf("HMV_CFI_CONEXIÓN");
            if (idx >= 0)
                cmbParameter.SelectedIndex = idx;
            else
                cmbParameter.Text = "HMV_CFI_CONEXIÓN";

            this.Closed += (s, e) => ElectricalGeometryCommand.ClearWindow();
        }

        // ── Called by the handler from any thread ──────────────────────
        public void SetStatus(string msg)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                txtStatus.Text += msg + Environment.NewLine;
                txtStatus.ScrollToEnd();
            }));
        }

        // ── UI events ──────────────────────────────────────────────────
        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void ChkIncludeDxf_Changed(object sender, RoutedEventArgs e)
        {
            pnlDxf.Visibility = chkIncludeDxf.IsChecked == true
                ? Visibility.Visible
                : Visibility.Collapsed;
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description      = "Select the folder containing the DXF files";
                dialog.ShowNewFolderButton = false;
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    txtFolder.Text = dialog.SelectedPath;
            }
        }

        private void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {
            cmbParameter.IsDropDownOpen = false;

            string paramText = cmbParameter.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(paramText))
                paramText = cmbParameter.Text;

            if (string.IsNullOrWhiteSpace(paramText))
            {
                SetStatus("Error: Please select or enter a Parameter Name.");
                return;
            }

            bool includeDxf = chkIncludeDxf.IsChecked == true;
            if (includeDxf && string.IsNullOrWhiteSpace(txtFolder.Text))
            {
                SetStatus("Error: Please select a DXF folder.");
                return;
            }

            // Ask for save path HERE on the UI thread — avoids Dispatcher deadlock in handler
            string outputPath = null;
            using (var dlg = new System.Windows.Forms.SaveFileDialog())
            {
                dlg.Filter   = "CSV Files (*.csv)|*.csv";
                dlg.Title    = "Save Electrical Geometry Report";
                dlg.FileName = "ElectricalGeometryReport.csv";

                if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                {
                    SetStatus("Save cancelled.");
                    return;
                }
                outputPath = dlg.FileName;
            }

            // Pass all settings to handler then raise the external event
            _handler.TargetParam = paramText;
            _handler.IncludeDxf  = includeDxf;
            _handler.DxfFolder   = includeDxf ? txtFolder.Text : null;
            _handler.OutputPath  = outputPath;

            SetStatus("Running...");
            _exEvent.Raise();
        }
    }
}
