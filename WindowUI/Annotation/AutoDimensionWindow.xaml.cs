using System;
using System.Windows;
using System.Windows.Input;

namespace HMVTools
{
    public partial class AutoDimensionWindow : Window
    {
        public double ToleranceCm { get; private set; }

        public AutoDimensionWindow()
        {
            InitializeComponent();
            txtTolerance.Focus();
        }

        // Draggable top bar
        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(txtTolerance.Text, out double tol))
            {
                ToleranceCm = tol;
                this.DialogResult = true; // Returns true to the Command execution
                this.Close();
            }
            else
            {
                MessageBox.Show("Please enter a valid number for tolerance.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}