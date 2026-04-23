using System.Windows;
using System.Windows.Input;
using System.Windows.Forms; // Required for FolderBrowserDialog
using Application = System.Windows.Application;

namespace HMVTools
{
    public partial class ElectricalGeometryWindow : Window
    {
        public string SelectedFolder { get; private set; }
        public string ParameterName { get; private set; }
        public bool Proceed { get; private set; } = false;

        public ElectricalGeometryWindow()
        {
            InitializeComponent();
        }

        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Proceed = false;
            this.Close();
        }

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select the folder containing the DXF files";
                dialog.ShowNewFolderButton = false;
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtFolder.Text = dialog.SelectedPath;
                }
            }
        }

        private void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtParameter.Text))
            {
                lblStatus.Text = "Error: Please enter a Parameter Name.";
                return;
            }
            if (string.IsNullOrWhiteSpace(txtFolder.Text))
            {
                lblStatus.Text = "Error: Please select a DXF folder.";
                return;
            }

            SelectedFolder = txtFolder.Text;
            ParameterName = txtParameter.Text;
            Proceed = true;
            this.Close();
        }
    }
}