using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HMVTools
{
    [Transaction(TransactionMode.ReadOnly)]
    public class LinkedFilesAuditCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            // 1. Collect all Link Instances
            var linkInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            if (!linkInstances.Any())
            {
                TaskDialog.Show("Audit", "No linked Revit files found in this project.");
                return Result.Cancelled;
            }

            List<LinkAuditData> auditResults = new List<LinkAuditData>();

            foreach (var instance in linkInstances)
            {
                var data = new LinkAuditData { LinkName = instance.Name };
                Document linkDoc = instance.GetLinkDocument();

                if (linkDoc != null)
                {
                    // Get Site Location (Lat/Long/GIS)
                    SiteLocation site = linkDoc.SiteLocation;

                    // FIXED: Using GeoCoordinateSystemId for Revit 2023
                    data.GisCode = string.IsNullOrEmpty(site.GeoCoordinateSystemId) ? "Default" : site.GeoCoordinateSystemId;

                    // Converting Radians to Degrees for Lat/Long
                    data.Latitude = (site.Latitude * 180 / Math.PI).ToString("F6");
                    data.Longitude = (site.Longitude * 180 / Math.PI).ToString("F6");

                    // Get Active Site Name
                    data.SiteName = linkDoc.ActiveProjectLocation.Name;

                    // Get Angle to True North
                    ProjectPosition position = linkDoc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero);
                    double angleDegrees = position.Angle * 180 / Math.PI;
                    data.TrueNorthAngle = angleDegrees.ToString("F2") + "°";
                }
                else
                {
                    data.GisCode = "N/A";
                    data.Latitude = "N/A";
                    data.Longitude = "N/A";
                    data.SiteName = "UNLOADED";
                    data.TrueNorthAngle = "N/A";
                }

                auditResults.Add(data);
            }

            // 2. Open the customized WPF UI
            var window = new LinkedFilesAuditWindow(auditResults);

            // Ensures the WPF window behaves properly as a child of the Revit main window
            System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(window);
            helper.Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;

            window.ShowDialog();

            return Result.Succeeded;
        }
    }

    // Data Class
    public class LinkAuditData
    {
        public string LinkName { get; set; }
        public string GisCode { get; set; }
        public string Latitude { get; set; }
        public string Longitude { get; set; }
        public string SiteName { get; set; }
        public string TrueNorthAngle { get; set; }
    }
}