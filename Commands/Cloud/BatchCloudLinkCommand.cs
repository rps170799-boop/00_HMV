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
                    // ─── PHASE 4: Discovery ──────────────────────────────────
                    // Step 1: Get hub
                    var hub = System.Threading.Tasks.Task.Run(async () =>
                    {
                        return await aps.GetHubAsync();
                    }).GetAwaiter().GetResult();

                    // Step 2: Get projects
                    var projects = System.Threading.Tasks.Task.Run(async () =>
                    {
                        return await aps.GetProjectsAsync(hub.Id);
                    }).GetAwaiter().GetResult();

                    if (projects.Count == 0)
                    {
                        TaskDialog.Show("HMV Tools", "No projects found in this hub.");
                        return Result.Cancelled;
                    }

                    // Step 3: Let user pick a project via TaskDialog CommandLinks
                    TaskDialog projectPicker = new TaskDialog("HMV Tools - Select Project");
                    projectPicker.MainInstruction = $"Hub: {hub.Name}";
                    projectPicker.MainContent = "Select the project to search for RVT files:";
                    projectPicker.CommonButtons = TaskDialogCommonButtons.Cancel;

                    // TaskDialog supports up to 4 command links (200-203)
                    // If more than 4 projects, we'll need a WPF dialog — but 5 fits with a trick:
                    // we use the "footer" approach. For now, limit to first 4.
                    int maxLinks = Math.Min(projects.Count, 4);
                    for (int i = 0; i < maxLinks; i++)
                    {
                        projectPicker.AddCommandLink((TaskDialogCommandLinkId)(200 + i), projects[i].Name);
                    }

                    TaskDialogResult projectChoice = projectPicker.Show();
                    int projectIndex = (int)projectChoice - 200;

                    if (projectIndex < 0 || projectIndex >= maxLinks)
                    {
                        return Result.Cancelled; // User clicked Cancel
                    }

                    string selectedProjectId = projects[projectIndex].Id;
                    string selectedProjectName = projects[projectIndex].Name;

                    // Step 4: Get top folders for the selected project
                    var folders = System.Threading.Tasks.Task.Run(async () =>
                    {
                        return await aps.GetTopFoldersAsync(hub.Id, selectedProjectId);
                    }).GetAwaiter().GetResult();

                    if (folders.Count == 0)
                    {
                        TaskDialog.Show("HMV Tools", $"No folders found in {selectedProjectName}.");
                        return Result.Cancelled;
                    }

                    // Step 5: Let user pick a top folder
                    TaskDialog folderPicker = new TaskDialog("HMV Tools - Select Folder");
                    folderPicker.MainInstruction = $"Project: {selectedProjectName}";
                    folderPicker.MainContent = "Select the folder to scan for RVT files:";
                    folderPicker.CommonButtons = TaskDialogCommonButtons.Cancel;

                    int maxFolderLinks = Math.Min(folders.Count, 4);
                    for (int i = 0; i < maxFolderLinks; i++)
                    {
                        folderPicker.AddCommandLink((TaskDialogCommandLinkId)(200 + i), folders[i].Name);
                    }

                    TaskDialogResult folderChoice = folderPicker.Show();
                    int folderIndex = (int)folderChoice - 200;

                    if (folderIndex < 0 || folderIndex >= maxFolderLinks)
                    {
                        return Result.Cancelled;
                    }

                    string selectedFolderId = folders[folderIndex].Id;
                    string selectedFolderName = folders[folderIndex].Name;

                    // Step 6: Optional — drill one level deeper into subfolders
                    var subFolders = System.Threading.Tasks.Task.Run(async () =>
                    {
                        return await aps.GetSubFoldersAsync(selectedProjectId, selectedFolderId);
                    }).GetAwaiter().GetResult();

                    if(subFolders.Count > 0)
                    {
                        TaskDialog subPicker = new TaskDialog ("HMV Tools - Select Subfolder (Optional)");
                        subPicker.MainInstruction = $"Folder: {selectedFolderName}";
                        subPicker.MainContent  = "Pick a subfolder to narrow the search,\nor click 'Scan entire folder' to scan everything.";
                        // First command link = scan entire folder (no drill-down)

                        subPicker.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                            $"Scan entire '{selectedFolderName}'",
                            $"Search all {subFolders.Count} subfolders recursively");

                        // Up to 3 subfolders as additional options

                        int maxSubs = Math.Min (subFolders.Count, 3);
                        for (int i=0; i < maxSubs; i++)
                        {
                            subPicker.AddCommandLink(
                                (TaskDialogCommandLinkId)(202 + i),
                                subFolders[i].Name);
                        }
                        subPicker.CommonButtons = TaskDialogCommonButtons.Cancel;
                        TaskDialogResult subChoice = subPicker.Show();
                        if (subChoice == TaskDialogResult.Cancel) return Result.Cancelled;

                        int subIndex = (int)subChoice - 202;

                        if(subIndex >=0 && subIndex < maxSubs)
                        {
                            // User picked a specific subfolder to scan
                            selectedFolderName = selectedFolderName + " / " + subFolders[subIndex].Name;
                            selectedFolderId = subFolders[subIndex].Id;
                        }
                        // else: user clicked "Scan entire folder" (CommandLink1) — keep original folder

                    }
                    // Step 7: Crawl for .rvt files
                    var rvtFiles = System.Threading.Tasks.Task.Run(async () =>
                    {
                        return await aps.FindRevitFilesAsync(selectedProjectId, selectedFolderId, selectedFolderName);
                    }).GetAwaiter().GetResult();

                    if( rvtFiles.Count == 0)
                    {
                        TaskDialog.Show("HMV Tools", $"No .rvt files found in '{selectedFolderName}'.");
                        return Result.Cancelled;
                    }
                    // Step 8: Show WPF window

                    var window = new BatchCloudLinkWindow(rvtFiles,selectedProjectName,selectedFolderName);
                    window.ShowDialog();
                    if (window.DialogResult != true || window.SelectedFiles == null || window.SelectedFiles.Count == 0)
                    {
                        return Result.Cancelled;
                    }
                    // Step 9: Placeholder — Phase 7 will do the actual linking here
                    var summary = new System.Text.StringBuilder();
                    summary.AppendLine($"Ready to link {window.SelectedFiles.Count} file(s):\n");
                    foreach (var f in window.SelectedFiles)
                    {
                        summary.AppendLine($"  • {f.Name} [{f.Placement}]");
                    }
                    summary.AppendLine($"\nPhase 7 will create the actual Revit links.");

                    TaskDialog.Show("HMV Tools - Link Preview", summary.ToString());

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