using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;

namespace HMVTools
{
    public partial class ElectricalFlowConfigWindow : Window
    {
        public ElectricalFlowConfig Result { get; private set; }

        public ElectricalFlowConfigWindow(
            List<string> equipParams,
            List<string> pointParams,
            ElectricalFlowConfig current)
        {
            InitializeComponent();

            cmbSourceEquipParam.ItemsSource = equipParams;
            cmbDestPointParam.ItemsSource   = pointParams;

            if (current != null)
            {
                SelectOrSet(cmbSourceEquipParam, current.SourceEquipmentParam);
                SelectOrSet(cmbDestPointParam,   current.DestPointParam);
            }
            else
            {
                // Try to select from the list; if the param isn't there yet, set as typed text.
                // This ensures the defaults are always visible even in an empty model.
                TrySelectDefault(cmbSourceEquipParam, "HMV_CARGAS DE DISEÑO_EQUIPO");
                TrySelectDefault(cmbDestPointParam,   "HMV_CFI_EQUIPO INICIAL");
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            Result = new ElectricalFlowConfig
            {
                SourceEquipmentParam = cmbSourceEquipParam.SelectedItem as string ?? cmbSourceEquipParam.Text,
                DestPointParam       = cmbDestPointParam.SelectedItem   as string ?? cmbDestPointParam.Text
            };

            DialogResult = true;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }

        private static void SelectOrSet(System.Windows.Controls.ComboBox cmb, string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return;
            int idx = cmb.Items.IndexOf(value);
            if (idx >= 0) cmb.SelectedIndex = idx;
            else          cmb.Text          = value;
        }

        private static void TrySelectDefault(System.Windows.Controls.ComboBox cmb, string defaultValue)
        {
            int idx = cmb.Items.IndexOf(defaultValue);
            if (idx >= 0)
                cmb.SelectedIndex = idx;
            else
                cmb.Text = defaultValue;  // always show the default, even if list is empty
        }
    }
}
