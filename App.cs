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

            // ── Annotation Tools Panel ──
            RibbonPanel panelAnnot = app.CreateRibbonPanel("HMV Tools", "Annotation Tools");

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
            PushButton btnTopo = panelAnnot.AddItem(btnTopoLines) as PushButton;
            BitmapImage iconTopo = LoadImage("HMVTools.Resources.topo_32.png");
            if (iconTopo != null) btnTopo.LargeImage = iconTopo;

            PushButtonData btnPipeAnnot = new PushButtonData(
                "PipeandFramingAnnotations",
                "Pipe, Framing\nAnnotations",
                path,
                "HMVTools.PlaceAnnotationsAlongPipeCmd");
            btnPipeAnnot.ToolTip = "Place annotations along paths";
            btnPipeAnnot.LongDescription =
                "Coloca instancias de una familia de anotación genérica a lo largo "
                + "del recorrido de tuberías (Pipe) , tuberías flexibles (FlexPipe) y tambien vigas (Structural Framing).\n\n"
                + "Funciones:\n"
                + "• Selección de Pipes , FlexPipes, y Structural Framing\n"
                + "• Espaciado configurable entre anotaciones\n"
                + "• Orientación automática según la tangente del recorrido\n"
                + "• Funciona con tramos rectos, arcos y splines";
            PushButton btnPA = panelAnnot.AddItem(btnPipeAnnot) as PushButton;
            BitmapImage iconPipeAnnot = LoadImage("HMVTools.Resources.pipeannotation_32.png");
            if (iconPipeAnnot != null) btnPA.LargeImage = iconPipeAnnot;

            PushButtonData btnGridLevel = new PushButtonData(
                "GridLevelExtent",
                "Grid/Level\nExtent",
                path,
                "HMVTools.GridLevelExtentCommand");
            btnGridLevel.ToolTip = "Switch grids and levels between 2D and 3D extent";
            btnGridLevel.LongDescription =
                "Converts grids and levels between 2D (ViewSpecific) and 3D (Model) "
                + "datum extent across multiple views.\n\n"
                + "Features:\n"
                + "• Toggle between 2D (independent per view) and 3D (global)\n"
                + "• Choose to process Grids, Levels, or both\n"
                + "• Select individual views or apply to all views at once\n"
                + "• Summary report of all changes made";
            PushButton btnGL = panelAnnot.AddItem(btnGridLevel) as PushButton;
            BitmapImage iconGL = LoadImage("HMVTools.Resources.gridlevel_32.png");
            if (iconGL != null) btnGL.LargeImage = iconGL;

            // ── Audit Panel ──
            RibbonPanel panelAudit = app.CreateRibbonPanel("HMV Tools", "Audit");

            PushButtonData btnTextAudit = new PushButtonData(
                "TextAudit",
                "Text\nAudit",
                path,
                "HMVTools.TextAuditCommand");
            btnTextAudit.ToolTip = "Standardize all text styles and annotation tag families";
            btnTextAudit.LongDescription =
                "Ejecuta un proceso completo de estandarización de textos en el proyecto:\n\n"
                + "Step 1 – Standardize Properties: normaliza Font (Arial), Width (1.0), "
                + "Bold (No), Background (Opaque) y Size (snap a 1.5/2/2.5/3/3.5 mm) "
                + "en todos los TextNoteTypes y DimensionTypes.\n\n"
                + "Step 2 – Merge Types: agrupa TextNoteTypes por tamaño, renombra el "
                + "keeper a HMV_General_Xmm Arial y reasigna todas las instancias.\n\n"
                + "Step 3 – Purge: elimina los tipos vacíos que quedaron sin instancias.\n\n"
                + "Step 4 – Tag Families: abre cada familia de anotación, estandariza "
                + "las propiedades de texto internas y recarga la familia al proyecto.\n\n"
                + "Al finalizar muestra un reporte completo de todos los cambios.";
            PushButton btnTA = panelAudit.AddItem(btnTextAudit) as PushButton;
            BitmapImage iconTA = LoadImage("HMVTools.Resources.textaudit_32.png");
            if (iconTA != null) btnTA.LargeImage = iconTA;

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