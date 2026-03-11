using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace HMVTools
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        {
            app.CreateRibbonTab("HMV Tools");
            string path = Assembly.GetExecutingAssembly().Location;

            // ── DWG Panel ──
            RibbonPanel panelDwg = app.CreateRibbonPanel("HMV Tools", "DWG");

            PushButtonData btnConvert = new PushButtonData(
                "DwgConvert",
                "DWG\nConvert",
                path,
                "HMVTools.DwgConvertCommand");
            btnConvert.ToolTip = "Convert DWG lines and texts to HMV standards";
            btnConvert.LongDescription =
                "Step 1 – Convert Lines: extracts geometry from the DWG import "
                + "and creates HMV_LINEA detail lines.\n"
                + "Step 2 – Partial Explode the DWG manually in Revit.\n"
                + "Step 3 – Standardize All: re-styles exploded lines to HMV_LINEA, "
                + "converts texts to HMV_General_<size> <font>, and purges "
                + "unused DWG-imported styles from the project.";
            PushButton btn = panelDwg.AddItem(btnConvert) as PushButton;
            BitmapImage icon = LoadImage("HMVTools.Resources.dwg_32.png");
            if (icon != null) btn.LargeImage = icon;

            // ── Family Control Tools Panel ──
            RibbonPanel panelFamily = app.CreateRibbonPanel("HMV Tools", "Family Control Tools");

            PushButtonData btnDuctos = new PushButtonData(
                "DuctosEditor",
                "Ductos\nEditor",
                path,
                "HMVTools.USDTuberiaConfigCommand");
            btnDuctos.ToolTip = "Configurador visual de tuberías USD";
            btnDuctos.LongDescription =
                "Seleccione una o más instancias de la familia USD y ejecute "
                + "este comando para editar visibilidad, diámetro y distancias "
                + "de las 45 tuberías (5×9) desde una interfaz visual.\n\n"
                + "Funciones:\n"
                + "• Grilla 5×9 con checkbox y diámetro por tubería\n"
                + "• Control de distancias X/Y entre ejes\n"
                + "• Herramientas masivas: toggle fila/columna, ø global\n"
                + "• Aplica cambios a todas las instancias seleccionadas.";
            PushButton btnD = panelFamily.AddItem(btnDuctos) as PushButton;
            BitmapImage iconDuctos = LoadImage("HMVTools.Resources.ductos_32.png");
            if (iconDuctos != null) btnD.LargeImage = iconDuctos;

            // ── Topography Panel ──
            RibbonPanel panelTopo = app.CreateRibbonPanel("HMV Tools", "Topography");

            PushButtonData btnTopoLines = new PushButtonData(
                "TopoToLines",
                "Topo to\nLines",
                path,
                "HMVTools.TopographyToLinesCommand");
            btnTopoLines.ToolTip = "Convert topography to detail lines";
            btnTopoLines.LongDescription =
                "Extrae las curvas de nivel de la topografía contenida en un vínculo "
                + "de Revit y las convierte en líneas de detalle nativas en la vista activa.\n\n"
                + "Funciones:\n"
                + "• Selección de vínculo de Revit con topografía\n"
                + "• Cálculo automático de intersecciones con el plano de vista\n"
                + "• Creación de líneas de detalle en la vista activa\n"
                + "• Reporte de líneas creadas y omitidas";
            PushButton btnTopo = panelTopo.AddItem(btnTopoLines) as PushButton;
            BitmapImage iconTopo = LoadImage("HMVTools.Resources.topo_32.png");
            if (iconTopo != null) btnTopo.LargeImage = iconTopo;

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }

        private BitmapImage LoadImage(string resourceName)
        {
            try
            {
                Stream stream = Assembly.GetExecutingAssembly()
                    .GetManifestResourceStream(resourceName);
                if (stream == null) return null;
                BitmapImage image = new BitmapImage();
                image.BeginInit();
                image.StreamSource = stream;
                image.EndInit();
                return image;
            }
            catch { return null; }
        }
    }
}