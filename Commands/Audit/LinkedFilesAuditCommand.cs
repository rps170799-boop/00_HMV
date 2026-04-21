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
                    // --- THE FIXED DESCRIPTION LOGIC ---
                    Parameter descParam = linkDoc.ProjectInformation.LookupParameter("BuildingDescription");
                    if (descParam != null && descParam.HasValue)
                    {
                        data.BuildingDescription = descParam.AsString();
                    }
                    else
                    {
                        data.BuildingDescription = "N/A";
                    }

                    // --- YOUR ORIGINAL WORKING BUILDING NAME CODE ---
                    string bName = linkDoc.ProjectInformation.BuildingName;
                    data.BuildingName = string.IsNullOrEmpty(bName) ? "N/A" : bName;


                    // Get Site Location (Lat/Long/GIS)
                    SiteLocation site = linkDoc.SiteLocation;

                    // FIXED: Using GeoCoordinateSystemId for Revit 2023
                    data.GisCode = string.IsNullOrEmpty(site.GeoCoordinateSystemId) ? "Default" : site.GeoCoordinateSystemId;

                    // Converting Radians to Degrees for Lat/Long
                    data.Latitude = (site.Latitude * 180 / Math.PI).ToString("F10");
                    data.Longitude = (site.Longitude * 180 / Math.PI).ToString("F10");

                    // Get Active Site Name
                    data.SiteName = linkDoc.ActiveProjectLocation.Name;

                    // Get Angle to True North
                    ProjectPosition position = linkDoc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero);
                    double angleDegrees = position.Angle * 180 / Math.PI;
                    data.TrueNorthAngle = angleDegrees.ToString("F2") + "°";

                    Element pbp = new FilteredElementCollector(linkDoc)
                        .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                        .WhereElementIsNotElementType()
                        .FirstOrDefault();
                    if (pbp != null)
                    {
                        string ns = SafeParamString(pbp, BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM);
                        string ew = SafeParamString(pbp, BuiltInParameter.BASEPOINT_EASTWEST_PARAM);
                        string elev = SafeParamString(pbp, BuiltInParameter.BASEPOINT_ELEVATION_PARAM);

                        // Fallback if all params came back N/A
                        if (ns == "N/A" && ew == "N/A" && elev == "N/A")
                        {
                            ProjectPosition pp = linkDoc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero);
                            double nsM = pp.NorthSouth * 0.3048;
                            double ewM = pp.EastWest * 0.3048;
                            double elevM = pp.Elevation * 0.3048;
                            data.ProjectBasePoint = $"N/S: {nsM:F10} m, E/W: {ewM:F10} m, Elev: {elevM:F10} m";
                        }
                        else
                        {
                            data.ProjectBasePoint = $"N/S: {ns}, E/W: {ew}, Elev: {elev}";
                        }
                    }

                    Element sp = new FilteredElementCollector(linkDoc)
                        .OfCategory(BuiltInCategory.OST_SharedBasePoint)
                        .WhereElementIsNotElementType()
                        .FirstOrDefault();
                    if (sp != null)
                    {
                        string ns = SafeParamString(sp, BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM);
                        string ew = SafeParamString(sp, BuiltInParameter.BASEPOINT_EASTWEST_PARAM);
                        string elev = SafeParamString(sp, BuiltInParameter.BASEPOINT_ELEVATION_PARAM);
                        data.SurveyPoint = $"N/S: {ns}, E/W: {ew}, Elev: {elev}";
                    }

                }
                else
                {
                    data.BuildingDescription = "UNLOADED";
                    data.BuildingName = "UNLOADED";
                    data.GisCode = "N/A";
                    data.Latitude = "N/A";
                    data.Longitude = "N/A";
                    data.SiteName = "UNLOADED";
                    data.TrueNorthAngle = "N/A";
                    data.ProjectBasePoint = "N/A";
                    data.SurveyPoint = "N/A";
                }

                auditResults.Add(data);
            }

            string hostName = doc.Title;
            if (hostName.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
            {
                hostName = hostName.Substring(0, hostName.Length - 4);
            }
            // --- ADD HOST MODEL DATA AT THE END ---
            var hostData = new LinkAuditData { LinkName = "<Host Model>: " + hostName + " ---" };

            Parameter hostDescParam = doc.ProjectInformation.LookupParameter("BuildingDescription");
            hostData.BuildingDescription = (hostDescParam != null && hostDescParam.HasValue) ? hostDescParam.AsString() : "N/A";

            string hostBName = doc.ProjectInformation.BuildingName;
            hostData.BuildingName = string.IsNullOrEmpty(hostBName) ? "N/A" : hostBName;

            SiteLocation hostSite = doc.SiteLocation;
            hostData.GisCode = string.IsNullOrEmpty(hostSite.GeoCoordinateSystemId) ? "Default" : hostSite.GeoCoordinateSystemId;
            hostData.Latitude = (hostSite.Latitude * 180 / Math.PI).ToString("F10");
            hostData.Longitude = (hostSite.Longitude * 180 / Math.PI).ToString("F10");
            hostData.SiteName = doc.ActiveProjectLocation.Name;

            ProjectPosition hostPosition = doc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero);
            hostData.TrueNorthAngle = (hostPosition.Angle * 180 / Math.PI).ToString("F2") + "°";

            Element hostPbp = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                .WhereElementIsNotElementType()
                .FirstOrDefault();
            if (hostPbp != null)
            {
                string ns = SafeParamString(hostPbp, BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM);
                string ew = SafeParamString(hostPbp, BuiltInParameter.BASEPOINT_EASTWEST_PARAM);
                string elev = SafeParamString(hostPbp, BuiltInParameter.BASEPOINT_ELEVATION_PARAM);

                if (ns == "N/A" && ew == "N/A" && elev == "N/A")
                {
                    double nsM = hostPosition.NorthSouth * 0.3048;
                    double ewM = hostPosition.EastWest * 0.3048;
                    double elevM = hostPosition.Elevation * 0.3048;
                    hostData.ProjectBasePoint = $"N/S: {nsM:F10} m, E/W: {ewM:F10} m, Elev: {elevM:F10} m";
                }
                else
                {
                    hostData.ProjectBasePoint = $"N/S: {ns}, E/W: {ew}, Elev: {elev}";
                }
            }
            else { hostData.ProjectBasePoint = "N/A"; }

            Element hostSp = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_SharedBasePoint)
                .WhereElementIsNotElementType()
                .FirstOrDefault();
            if (hostSp != null)
            {
                string ns = SafeParamString(hostSp, BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM);
                string ew = SafeParamString(hostSp, BuiltInParameter.BASEPOINT_EASTWEST_PARAM);
                string elev = SafeParamString(hostSp, BuiltInParameter.BASEPOINT_ELEVATION_PARAM);
                hostData.SurveyPoint = $"N/S: {ns}, E/W: {ew}, Elev: {elev}";
            }
            else { hostData.SurveyPoint = "N/A"; }

            auditResults.Add(hostData);
            // --- END HOST MODEL DATA ---


            // 2. Open the customized WPF UI
            var window = new LinkedFilesAuditWindow(auditResults, hostName);

            // Ensures the WPF window behaves properly as a child of the Revit main window
            System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(window);
            helper.Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;

            window.ShowDialog();

            return Result.Succeeded;
        }

        private string SafeParamString(Element el, BuiltInParameter bip)
        {
            if (el == null) return "N/A";
            Parameter p = el.get_Parameter(bip);
            if (p == null || !p.HasValue) return "N/A";

            // ALWAYS read as double for coordinate params to preserve full precision
            if (p.StorageType == StorageType.Double)
            {
                double feet = p.AsDouble();
                double meters = feet * 0.3048;
                return meters.ToString("R", System.Globalization.CultureInfo.InvariantCulture) + " m";
            }

            // Fallback for non-double params
            string vs = p.AsValueString();
            return !string.IsNullOrEmpty(vs) ? vs : "N/A";
        }
    }

    // Data Class (Property order correctly updated here too!)
    public class LinkAuditData
    {
        public string LinkName { get; set; }
        public string BuildingDescription { get; set; }
        public string BuildingName { get; set; }
        public string GisCode { get; set; }
        public string Latitude { get; set; }
        public string Longitude { get; set; }
        public string SiteName { get; set; }
        public string TrueNorthAngle { get; set; }
        public string ProjectBasePoint { get; set; }
        public string SurveyPoint { get; set; }
    }
}