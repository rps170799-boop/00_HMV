using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using System.Windows.Forms;

namespace HMVTools
{
    public partial class ElectricalRefresherConfigWindow : Window
    {
        public ElectricalRefresherConfig Result { get; private set; }

        // ── Constructor ────────────────────────────────────────────────────────
        public ElectricalRefresherConfigWindow(
            List<string>              pointParams,
            List<string>              flexPipeParams,
            List<string>              flexPipeTypes,
            List<string>              systemTypes,
            ElectricalRefresherConfig current)
        {
            InitializeComponent();

            // Adaptive point dropdowns (1-3)
            cmbCnxNumber.ItemsSource    = pointParams;
            cmbEquipoInicial.ItemsSource = pointParams;
            cmbConnected.ItemsSource    = pointParams;

            // FlexPipe dropdowns (4-6)
            cmbFlexCnxNumber.ItemsSource = flexPipeParams;
            cmbFlexInicial.ItemsSource   = flexPipeParams;
            cmbFlexFinal.ItemsSource     = flexPipeParams;

            // MEP type dropdowns (7-8)
            cmbFlexPipeType.ItemsSource = flexPipeTypes;
            cmbSystemType.ItemsSource   = systemTypes;

            if (current != null)
            {
                SelectOrSet(cmbCnxNumber,     current.CnxNumberParam);
                SelectOrSet(cmbEquipoInicial,  current.EquipoInicialParam);
                SelectOrSet(cmbConnected,      current.ConnectedParam);
                SelectOrSet(cmbFlexCnxNumber,  current.FlexCnxNumberParam);
                SelectOrSet(cmbFlexInicial,    current.FlexEquipoInicialParam);
                SelectOrSet(cmbFlexFinal,      current.FlexEquipoFinalParam);
                SelectOrSet(cmbFlexPipeType,   current.FlexPipeTypeKey);
                SelectOrSet(cmbSystemType,     current.ElectricalSystemTypeKey);

                if (!string.IsNullOrWhiteSpace(current.DxfFolder))
                    txtDxfFolder.Text = current.DxfFolder;
            }
            else
            {
                // Pre-select defaults
                TrySelectDefault(cmbCnxNumber,     "HMV_CFI_CONEXI\u00d3N");
                TrySelectDefault(cmbEquipoInicial,  "HMV_CFI_EQUIPO INICIAL");
                TrySelectDefault(cmbConnected,      "HMV_CFI_CONNECTED");
                TrySelectDefault(cmbFlexCnxNumber,  "HMV_CFI_CONEXI\u00d3N");
                TrySelectDefault(cmbFlexInicial,    "HMV_CFI_EQUIPO INICIAL");
                TrySelectDefault(cmbFlexFinal,      "HMV_CFI_EQUIPO FINAL");
            }
        }

        // ── Button handlers ────────────────────────────────────────────────────

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtDxfFolder.Text))
            {
                System.Windows.MessageBox.Show("Please select a DXF folder before saving.",
                    "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Result = new ElectricalRefresherConfig
            {
                CnxNumberParam         = cmbCnxNumber.SelectedItem    as string ?? cmbCnxNumber.Text,
                EquipoInicialParam     = cmbEquipoInicial.SelectedItem as string ?? cmbEquipoInicial.Text,
                ConnectedParam         = cmbConnected.SelectedItem     as string ?? cmbConnected.Text,
                FlexCnxNumberParam     = cmbFlexCnxNumber.SelectedItem as string ?? cmbFlexCnxNumber.Text,
                FlexEquipoInicialParam = cmbFlexInicial.SelectedItem   as string ?? cmbFlexInicial.Text,
                FlexEquipoFinalParam   = cmbFlexFinal.SelectedItem     as string ?? cmbFlexFinal.Text,
                FlexPipeTypeKey        = cmbFlexPipeType.SelectedItem  as string ?? cmbFlexPipeType.Text,
                ElectricalSystemTypeKey = cmbSystemType.SelectedItem   as string ?? cmbSystemType.Text,
                DxfFolder              = txtDxfFolder.Text.Trim()
            };

            DialogResult = true;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select the folder containing the DXF files";
                dlg.ShowNewFolderButton = false;
                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    txtDxfFolder.Text = dlg.SelectedPath;
            }
        }

        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => this.DragMove();

        // ── Helpers ────────────────────────────────────────────────────────────

        private static void SelectOrSet(System.Windows.Controls.ComboBox cmb, string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return;
            int idx = cmb.Items.IndexOf(value);
            if (idx >= 0) cmb.SelectedIndex = idx;
            else          cmb.Text          = value;
        }

        private static void TrySelectDefault(System.Windows.Controls.ComboBox cmb, string defaultValue)
        {
            if (cmb.SelectedItem != null) return;
            int idx = cmb.Items.IndexOf(defaultValue);
            if (idx >= 0) cmb.SelectedIndex = idx;
            else if (cmb.Items.Count > 0) cmb.SelectedIndex = 0;
        }
    }
}
