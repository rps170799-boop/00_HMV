using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace HMVTools
{
    // Clase contenedora de datos
    public class GridLevelViewItem
    {
        public int Index { get; set; }
        public string DisplayName { get; set; }
    }

    public partial class GridLevelExtentWindow : Window
    {
        private List<GridLevelViewItem> allItems;
        private List<bool> checkedState;


        private int lastClickedIndex = -1;

        // Propiedades públicas para que el Comando las lea al finalizar
        public bool ConvertTo2D { get; private set; }
        public bool ProcessGrids { get; private set; }
        public bool ProcessLevels { get; private set; }
        public List<int> SelectedIndices { get; private set; }

        public GridLevelExtentWindow(List<GridLevelViewItem> items)
        {
            InitializeComponent();

            allItems = items;
            checkedState = new List<bool>(new bool[items.Count]);

            PopulateViewList();
            UpdateCount();

            this.Loaded += (s, e) => searchBox.Focus();
        }

        // ─── LÓGICA DE LA BARRA SUPERIOR PERSONALIZADA ───

        // Permite mover la ventana al arrastrar la barra superior
        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        // Botón "X" de arriba a la derecha
        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        // Botón "Cancel" de abajo
        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        // ─── LÓGICA DE LA INTERFAZ ───

        private void PopulateViewList()
        {
            viewList.Items.Clear();
            string filter = searchBox.Text.ToLower();

            // Reseteamos la memoria si el usuario usa la barra de búsqueda
            lastClickedIndex = -1;

            for (int i = 0; i < allItems.Count; i++)
            {
                if (!string.IsNullOrEmpty(filter) && !allItems[i].DisplayName.ToLower().Contains(filter))
                    continue;

                int idx = i;
                var cb = new CheckBox
                {
                    Content = allItems[i].DisplayName,
                    IsChecked = checkedState[i],
                    FontSize = 10,
                    Padding = new Thickness(4),

                    // 👇 Esto alinea el texto perfectamente con el cuadrito:
                    VerticalContentAlignment = VerticalAlignment.Center,
                    Cursor = Cursors.Hand
                };

                // 👇 Lógica para habilitar la selección múltiple con Shift
                cb.Click += (s, e) =>
                {
                    int currentIndex = viewList.Items.IndexOf(cb);
                    bool isShiftPressed = Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift);
                    bool targetState = cb.IsChecked == true;

                    if (isShiftPressed && lastClickedIndex != -1)
                    {
                        int start = Math.Min(lastClickedIndex, currentIndex);
                        int end = Math.Max(lastClickedIndex, currentIndex);

                        for (int j = start; j <= end; j++)
                        {
                            if (viewList.Items[j] is CheckBox listCb)
                            {
                                listCb.IsChecked = targetState; // Aplica el cambio a todos los de en medio
                            }
                        }
                    }
                    lastClickedIndex = currentIndex; // Guarda el último clic
                };

                // Tus eventos originales que actualizan el conteo
                cb.Checked += (s, e) => { checkedState[idx] = true; UpdateCount(); };
                cb.Unchecked += (s, e) => { checkedState[idx] = false; UpdateCount(); };

                viewList.Items.Add(cb);
            }
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            PopulateViewList();
        }

        private void SelectAllViews_Click(object sender, RoutedEventArgs e)
        {
            bool allChecked = checkedState.All(c => c);
            for (int i = 0; i < checkedState.Count; i++)
            {
                checkedState[i] = !allChecked;
            }
            PopulateViewList();
            UpdateCount();
        }

        private void UpdateCount()
        {
            int sel = checkedState.Count(c => c);
            countText.Text = $"{sel} / {allItems.Count} selected";
        }

        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            ConvertTo2D = radio2D.IsChecked == true;
            ProcessGrids = chkGrids.IsChecked == true;
            ProcessLevels = chkLevels.IsChecked == true;

            SelectedIndices = new List<int>();
            for (int i = 0; i < allItems.Count; i++)
            {
                if (checkedState[i])
                    SelectedIndices.Add(allItems[i].Index);
            }

            if (!ProcessGrids && !ProcessLevels)
            {
                MessageBox.Show("Select at least Grids, Levels, or both.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (SelectedIndices.Count == 0)
            {
                MessageBox.Show("Select at least one view.", "HMV Tools", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            this.DialogResult = true;
            this.Close();
        }
    }
}