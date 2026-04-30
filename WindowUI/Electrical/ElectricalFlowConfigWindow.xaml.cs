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
            cmbDestEquipParam.ItemsSource   = pointParams;
            cmbDestCNParam.ItemsSource      = pointParams;

            // Pre-select values from the current config
            if (current != null)
            {
                SelectOrSet(cmbSourceEquipParam, current.SourceEquipmentParam);
                SelectOrSet(cmbDestEquipParam,   current.DestEquipmentNameParam);
                SelectOrSet(cmbDestCNParam,      current.DestCNParam);
                chkUpdateEquipNames.IsChecked = current.UpdateEquipmentNames;
                chkUpdateCN.IsChecked         = current.UpdateCN;
            }
            else
            {
                if (equipParams.Count > 0) cmbSourceEquipParam.SelectedIndex = 0;
                if (pointParams.Count > 0) cmbDestEquipParam.SelectedIndex   = 0;
                if (pointParams.Count > 1) cmbDestCNParam.SelectedIndex      = 1;
                else if (pointParams.Count > 0) cmbDestCNParam.SelectedIndex = 0;
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            Result = new ElectricalFlowConfig
            {
                SourceEquipmentParam   = cmbSourceEquipParam.SelectedItem as string ?? cmbSourceEquipParam.Text,
                DestEquipmentNameParam = cmbDestEquipParam.SelectedItem  as string ?? cmbDestEquipParam.Text,
                DestCNParam            = cmbDestCNParam.SelectedItem     as string ?? cmbDestCNParam.Text,
                UpdateEquipmentNames   = chkUpdateEquipNames.IsChecked == true,
                UpdateCN               = chkUpdateCN.IsChecked          == true
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
    }
}
