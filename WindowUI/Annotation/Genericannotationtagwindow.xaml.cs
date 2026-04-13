using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace HMVTools
{
    // NO CAMBIAR: El Command.cs depende de estas clases exactas
    public class TextTypeEntry
    {
        public string TypeName { get; set; }
        public int TypeIdInt { get; set; }
    }

    public class GenericAnnotationTagSettings
    {
        public int TextTypeIdInt { get; set; }
        public double OffsetXMm { get; set; }
        public double OffsetYMm { get; set; }
        public string ParameterName { get; set; }
    }

    public partial class GenericAnnotationTagWindow : Window
    {
        private List<TextTypeEntry> allTypes;
        private List<string> allParameters;
        public GenericAnnotationTagSettings Settings { get; private set; }

        public GenericAnnotationTagWindow(List<TextTypeEntry> types, List<string> parameters)
        {
            InitializeComponent();
            this.allTypes = types;
            this.allParameters = parameters;

            // Población inicial
            typeListBox.ItemsSource = allTypes;
            paramListBox.ItemsSource = allParameters;
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

        private void TypeSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string filter = typeSearchBox.Text.ToLower();
            typeListBox.ItemsSource = allTypes
                .Where(t => t.TypeName.ToLower().Contains(filter))
                .ToList();
        }

        private void ParamSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string filter = paramSearchBox.Text.ToLower();
            paramListBox.ItemsSource = allParameters
                .Where(p => p.ToLower().Contains(filter))
                .ToList();
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            if (typeListBox.SelectedItem == null || paramListBox.SelectedItem == null)
            {
                MessageBox.Show("Please select both a Text Type and a Parameter.", "HMV Tools");
                return;
            }

            double ox, oy;
            if (!double.TryParse(offsetXBox.Text, out ox)) ox = 0;
            if (!double.TryParse(offsetYBox.Text, out oy)) oy = 0;

            Settings = new GenericAnnotationTagSettings
            {
                TextTypeIdInt = (typeListBox.SelectedItem as TextTypeEntry).TypeIdInt,
                ParameterName = paramListBox.SelectedItem.ToString(),
                OffsetXMm = ox,
                OffsetYMm = oy
            };

            DialogResult = true;
            Close();
        }
    }
}