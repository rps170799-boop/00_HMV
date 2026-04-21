using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using static HMVTools.ElectricalConnectionCommand;


namespace HMVTools
{
    public partial class ConexionFlexPipeWindow : Window
    {
        private UIApplication _uiapp;
        private ElectricalConnectionHandler _handler;
        private ExternalEvent _exEvent;

        // Propiedades de estado para la UI
        private ElementId _elementAId = null;
        private ElementId _elementBId = null;

        public ConexionFlexPipeWindow(UIApplication uiapp, ElectricalConnectionHandler handler, ExternalEvent exEvent)
        {
            InitializeComponent();
            _uiapp = uiapp;
            _handler = handler;
            _exEvent = exEvent;
            _handler.UI = this;

            // Suscribirse al evento de cierre para limpiar la referencia estática
            this.Closed += (s, e) => { ElectricalConnectionCommand.ClearWindow(); };

            LoadRevitTypes();
        }

        public void SetStatus(string message)
        {
            // Usar BeginInvoke es más seguro en eventos externos de Revit
            Dispatcher.BeginInvoke(new Action(() =>
            {
                txtStatus.Text = message;
            }));
        }

        private void LoadRevitTypes()
        {
            Document doc = _uiapp.ActiveUIDocument.Document;

            // Cargar tipos de FlexPipe
            var flexTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(FlexPipeType))
                .Cast<FlexPipeType>()
                .OrderBy(t => t.Name)
                .ToList();
            cmbFlexPipeType.ItemsSource = flexTypes;
            if (flexTypes.Any()) cmbFlexPipeType.SelectedIndex = 0;

            // Cargar tipos de Sistemas (MEPSystemType es la clase correcta para sistemas eléctricos, tuberías, etc.)
            var sysTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(MEPSystemType))
                .Cast<MEPSystemType>()
                // Opcional: Si quieres filtrar solo los eléctricos, puedes descomentar la siguiente línea:
                // .Where(t => t.SystemClassification == MEPSystemClassification.PowerCircuit || t.SystemClassification == MEPSystemClassification.DataCircuit)
                .OrderBy(t => t.Name)
                .ToList();

            cmbSystemType.ItemsSource = sysTypes;

            if (cmbSystemType.Items.Count > 0) cmbSystemType.SelectedIndex = 0;
            // Si tu plantilla no tiene ElectricalSystemType, podemos cargar MEPSystemType y filtrar
            if (!sysTypes.Any())
            {
                var mepSysTypes = new FilteredElementCollector(doc)
                    .OfClass(typeof(MEPSystemType))
                    .Cast<MEPSystemType>()
                    .OrderBy(t => t.Name)
                    .ToList();
                cmbSystemType.ItemsSource = mepSysTypes;
            }
            else
            {
                cmbSystemType.ItemsSource = sysTypes;
            }

            if (cmbSystemType.Items.Count > 0) cmbSystemType.SelectedIndex = 0;
        }

        // ─── Funciones de la Interfaz ──────────────────────────────────────────

        private void TopBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "DXF Files (*.dxf)|*.dxf",
                Title = "Seleccionar trayectoria DXF"
            };

            if (ofd.ShowDialog() == true)
            {
                txtDxfPath.Text = ofd.FileName;
            }
        }

        // ─── Selección de Elementos en Revit ──────────────────────────────────

        private void BtnPickA_Click(object sender, RoutedEventArgs e)
        {
            PickElement(ref _elementAId, lblElementA, "Element A (Start)");
        }

        private void BtnPickB_Click(object sender, RoutedEventArgs e)
        {
            PickElement(ref _elementBId, lblElementB, "Element B (End)");
        }

        private void PickElement(ref ElementId targetId, System.Windows.Controls.TextBlock label, string prefix)
        {
            this.Hide(); // Oculta la ventana WPF para permitir interacción con Revit
            try
            {
                Reference r = _uiapp.ActiveUIDocument.Selection.PickObject(
                    ObjectType.Element,
                    $"Seleccione el equipo para {prefix}");

                targetId = r.ElementId;
                Element elem = _uiapp.ActiveUIDocument.Document.GetElement(targetId);
                label.Text = $"{prefix}: {elem.Name}";
                label.Foreground = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#004982"));
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // El usuario presionó ESC, no hacemos nada.
            }
            finally
            {
                this.ShowDialog(); // Vuelve a mostrar la ventana
            }
        }

        // ─── Ejecución ─────────────────────────────────────────────────────────

        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            // Validaciones
            if (string.IsNullOrWhiteSpace(txtDxfPath.Text))
            {
                SetStatus("Error: Seleccione un archivo DXF.");
                return;
            }
            if (_elementAId == null || _elementBId == null)
            {
                SetStatus("Error: Seleccione ambos elementos (A y B).");
                return;
            }
            if (cmbFlexPipeType.SelectedItem == null || cmbSystemType.SelectedItem == null)
            {
                SetStatus("Error: Seleccione los tipos de FlexPipe y Sistema.");
                return;
            }
            if (string.IsNullOrWhiteSpace(txtParamName.Text))
            {
                SetStatus("Error: Indique el nombre del parámetro compartido.");
                return;
            }

            // Poblamos las propiedades del Handler antes de ejecutar
            _handler.DxfFilePath = txtDxfPath.Text;
            _handler.ElementAId = _elementAId;
            _handler.ElementBId = _elementBId;
            _handler.SharedParameterName = txtParamName.Text.Trim();

            // Pasamos los nombres (Keys) como lo programó Opus
            _handler.SelectedFlexPipeTypeKey = ((ElementType)cmbFlexPipeType.SelectedItem).Name;
            _handler.SelectedElectricalSystemTypeKey = ((ElementType)cmbSystemType.SelectedItem).Name;

            SetStatus("Procesando geometría DXF y generando FlexPipe...");

            // Disparamos el evento al Handler
            _exEvent.Raise();
        }

        
    }
}