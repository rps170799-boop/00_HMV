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
                if (equipParams.Count > 0) cmbSourceEquipParam.SelectedIndex = 0;
                if (pointParams.Count > 0) cmbDestPointParam.SelectedIndex   = 0;
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
    }
}
