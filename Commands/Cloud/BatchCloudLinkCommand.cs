using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class BatchCloudLinkCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            // 1. Get the active Revit application and document
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // 2. Safety Check: Are we actually in a cloud model?
            if (!doc.IsModelInCloud)
            {
                TaskDialog.Show("HMV Tools",
                    "You must be working inside an ACC Cloud Model to use the Batch Linker.");
                return Result.Cancelled;
            }

            try
            {
                // 3. The Auto-Detect Magic
                ModelPath cloudPath = doc.GetCloudModelPath();
                Guid revitProjectId = cloudPath.GetProjectGUID();

                // 4. The APS Gotcha: Add the "b." prefix!
                // Revit gives us a standard GUID, but APS requires BIM 360/ACC 
                // project IDs to start with "b."
                string apsProjectId = "b." + revitProjectId.ToString();

                // (Optional for testing: Let's prove we got it)
                TaskDialog.Show("HMV Tools - Success!",
                    $"Extracted APS Project ID:\n{apsProjectId}\n\nWe are ready to search the cloud!");

                // ═══════════════════════════════════════════════════
                // PHASE 2: PKCE 3-Legged Authentication
                ApsManager aps = new ApsManager();

                // CRITICAL FIX 1: Force modern web security (TLS 1.2)
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

                // CRITICAL FIX 2: Run the web request on a background thread to prevent Revit from deadlocking
                bool isAuthSuccessful = System.Threading.Tasks.Task.Run(async () =>
                {
                    return await aps.AuthenticateUserInteractiveAsync();
                }).GetAwaiter().GetResult();

                if (isAuthSuccessful)
                {

                    // Phase 2: confirm login worked before re-running the diagnostic
                    TaskDialog.Show("HMV Tools - Login Success",
                        $"Login successful!\n\nNow re-running the diagnostic with your USER token...\n\nProject ID: {apsProjectId}");

                    string diagnosticReport = System.Threading.Tasks.Task.Run(async () =>
                    {
                        return await aps.RunDiagnosticV4Async();
                    }).GetAwaiter().GetResult();

                    string logPath = System.IO.Path.Combine(
                        System.IO.Path.GetTempPath(),
                        $"HMV_APS_DiagnosticV2_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                    System.IO.File.WriteAllText(logPath, diagnosticReport);

                    TaskDialog td = new TaskDialog("HMV Tools - Phase 2 Diagnostic V4")
                    {
                        MainInstruction = "Diagnostic with USER token complete.",
                        MainContent = $"Full report saved to:\n{logPath}\n\nClick 'Show details' for inline view.",
                        ExpandedContent = diagnosticReport,
                        CommonButtons = TaskDialogCommonButtons.Close
                    };
                    td.Show();
                }
                else
                {
                    return Result.Failed;
                }
                // ═══════════════════════════════════════════════════
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}