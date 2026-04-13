using System.Windows;
using System.Windows.Input;

namespace HMVTools
{
    public partial class ConstraintsReleaseWindow : Window
    {
        public ConstraintsReleaseWindow(int selectedCount)
        {
            InitializeComponent();

            // Seteamos el texto informativo con el conteo
            txtCountInfo.Text = $"You have {selectedCount} element(s) selected.";
        }

        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void BtnRelease_Click(object sender, RoutedEventArgs e)
        {
            // Al devolver true, el Command.cs ejecutará la transacción de limpieza
            DialogResult = true;
            Close();
        }
    }
}