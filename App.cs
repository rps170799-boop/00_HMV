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

            // DWG Panel - single button
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