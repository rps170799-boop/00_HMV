using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HMVTools
{
    // ── Enums & Settings ───────────────────────────────────────
    public enum PlacementMode
    {
        GenericAnnotation,
        DetailItem
    }

    public class PipeAnnotationSettings
    {
        public PlacementMode Mode { get; set; }
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
        /// <summary>Only used in GenericAnnotation mode.</summary>
        public double SpacingMm { get; set; }
    }

    // ── Entry container ────────────────────────────────────────
    public class FamilyEntry
    {
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
        public string Display => $"{FamilyName} : {TypeName}";
    }

    // ── Window ─────────────────────────────────────────────────
    public partial class PipeAnnotationWindow : Window
    {
        private List<FamilyEntry> annotItems;
        private List<FamilyEntry> detailItems;
        private PlacementMode currentMode;

        // 👇 CAMBIO AQUÍ: Se renombró a "Settings" para coincidir con tu comando
        public PipeAnnotationSettings Settings { get; private set; }

        // 👇 CAMBIO AQUÍ: Se le agregó "= 2000.0" al final para que no dé error si el comando no lo envía
        public PipeAnnotationWindow(List<FamilyEntry> annotations, List<FamilyEntry> details, double defaultSpacing = 2000.0)
        {
            InitializeComponent();

            annotItems = annotations ?? new List<FamilyEntry>();
            detailItems = details ?? new List<FamilyEntry>();
            spacingBox.Text = defaultSpacing.ToString("F0");

            SetMode(PlacementMode.GenericAnnotation);

            this.Loaded += (s, e) => searchBox.Focus();
        }

        // ── Lógica de la Barra Superior ──
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

        // ── Botones Inferiores ──
        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            warningBorder.Visibility = Visibility.Collapsed;

            var sel = listBox.SelectedItem as FamilyEntry;
            if (sel == null)
            {
                ShowWarning("Please select a family type.");
                return;
            }

            double space = 0;
            if (currentMode == PlacementMode.GenericAnnotation)
            {
                if (!double.TryParse(spacingBox.Text, out space) || space <= 0)
                {
                    ShowWarning("Please enter a valid positive number for spacing.");
                    return;
                }
            }

            Settings = new PipeAnnotationSettings
            {
                Mode = currentMode,
                FamilyName = sel.FamilyName,
                TypeName = sel.TypeName,
                SpacingMm = space
            };

            this.DialogResult = true;
            this.Close();
        }

        // ── Cambio de Modos ──
        private void AnnotBtn_Click(object sender, MouseButtonEventArgs e)
        {
            SetMode(PlacementMode.GenericAnnotation);
        }

        private void DetailBtn_Click(object sender, MouseButtonEventArgs e)
        {
            SetMode(PlacementMode.DetailItem);
        }

        private void SetMode(PlacementMode mode)
        {
            currentMode = mode;
            searchBox.Text = "";

            var activeBg = new SolidColorBrush(Color.FromRgb(0, 73, 130)); 
            var inactiveBg = new SolidColorBrush(Color.FromRgb(240, 240, 243)); // #F0F0F3

            if (mode == PlacementMode.GenericAnnotation)
            {
                annotBtnBorder.Background = activeBg;
                annotBtnText.Foreground = Brushes.White;
                annotBtnText.FontWeight = FontWeights.SemiBold;

                detailBtnBorder.Background = inactiveBg;
                detailBtnText.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 60)); // #3C3C3C
                detailBtnText.FontWeight = FontWeights.Normal;

                spacingPanel.Visibility = Visibility.Visible;
            }
            else
            {
                detailBtnBorder.Background = activeBg;
                detailBtnText.Foreground = Brushes.White;
                detailBtnText.FontWeight = FontWeights.SemiBold;

                annotBtnBorder.Background = inactiveBg;
                annotBtnText.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 60)); // #3C3C3C
                annotBtnText.FontWeight = FontWeights.Normal;

                spacingPanel.Visibility = Visibility.Collapsed;
            }

            PopulateList();
        }

        // ── Lógica de Listado y Búsqueda ──
        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            PopulateList();
        }

        private void PopulateList()
        {
            listBox.Items.Clear();
            string filter = searchBox.Text.ToLower();

            var sourceList = (currentMode == PlacementMode.GenericAnnotation) ? annotItems : detailItems;

            foreach (var item in sourceList)
            {
                if (string.IsNullOrEmpty(filter) || item.Display.ToLower().Contains(filter))
                {
                    listBox.Items.Add(item);
                }
            }
            // Muestra "Display" porque ListBox llama al ToString() o mapea el nombre, pero para forzarlo usaremos DisplayMemberPath en el CS:
            listBox.DisplayMemberPath = "Display";
        }

        private void ShowWarning(string msg)
        {
            warningText.Text = msg;
            warningBorder.Visibility = Visibility.Visible;
        }
    }
}