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

            RibbonPanel panelDwg = app.CreateRibbonPanel("HMV Tools", "DWG");
            string path = Assembly.GetExecutingAssembly().Location;

            PushButtonData btnData = new PushButtonData(
                "DwgToLines",
                "DWG to\nLines",
                path,
                "HMVTools.DwgToLinesCommand");

            PushButton btn = panelDwg.AddItem(btnData) as PushButton;
            btn.ToolTip = "Convert DWG lines to standardized Revit detail lines";
            btn.LongDescription = "Reads line weights and patterns from an imported DWG "
                + "and creates HMV-standardized detail lines with matching line styles.";

            // Load icon if available
            BitmapImage icon = LoadImage("HMVTools.Resources.dwg_32.png");
            if (icon != null)
                btn.LargeImage = icon;

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
            catch
            {
                return null;
            }
        }
    }
}