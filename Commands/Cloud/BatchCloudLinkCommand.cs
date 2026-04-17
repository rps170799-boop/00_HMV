using System;
using System.Collections.Generic;
using System.Linq;
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
                        return await aps.FindRevitFilesAsync(selectedProjectId, selectedFolderId);
                    }).GetAwaiter().GetResult();

                    if (rvtFiles.Count == 0)
                    {
                        TaskDialog.Show("HMV Tools", $"No .rvt files found in '{selectedFolderName}'.");
                        return Result.Cancelled;
                    }

                    // Step 8: Detect existing links to avoid duplicates
                    var existingLinkNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    var linkCollector = new FilteredElementCollector(doc);
                    foreach (RevitLinkType linkType in linkCollector.OfClass(typeof(RevitLinkType)).Cast<RevitLinkType>())
                    {
                        string linkName = linkType.Name;
                        existingLinkNames.Add(linkName);
                        if (linkName.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
                            existingLinkNames.Add(linkName.Substring(0, linkName.Length - 4));
                    }

                    // Filter out already-linked files
                    var availableFiles = rvtFiles.Where(f =>
                        !existingLinkNames.Contains(f.Name) &&
                        !existingLinkNames.Contains(System.IO.Path.GetFileNameWithoutExtension(f.Name))
                    ).ToList();

                    int alreadyLinked = rvtFiles.Count - availableFiles.Count;

                    if (availableFiles.Count == 0)
                    {
                        TaskDialog.Show("HMV Tools",
                            $"All {rvtFiles.Count} RVT file(s) are already linked in this model.");
                        return Result.Cancelled;
                    }

                    // Step 9: Show WPF window
                    string windowSubtitle = $"{selectedProjectName} / {selectedFolderName}";
                    if (alreadyLinked > 0)
                        windowSubtitle += $"  ({alreadyLinked} already linked — excluded)";

                    var window = new BatchCloudLinkWindow(availableFiles, selectedProjectName, selectedFolderName);
                    if (alreadyLinked > 0)
                        window.SubHeaderLabel.Text = windowSubtitle;

                    window.ShowDialog();

                    if (window.DialogResult != true || window.SelectedFiles == null || window.SelectedFiles.Count == 0)
                        return Result.Cancelled;

                    // Step 10: Resolve cloud GUIDs and create links
                    var results = new System.Text.StringBuilder();
                    int successCount = 0;
                    int failCount = 0;

                    foreach (var file in window.SelectedFiles)
                    {
                        try
                        {
                            // Resolve model GUID from APS (background thread)
                            var guids = System.Threading.Tasks.Task.Run(async () =>
                            {
                                return await aps.ResolveCloudGuidsAsync(selectedProjectId, file.Urn);
                            }).GetAwaiter().GetResult();

                            if (guids == null)
                            {
                                results.AppendLine($"FAILED: {file.Name} — Could not resolve cloud GUIDs");
                                failCount++;
                                continue;
                            }

                            Guid linkProjectGuid = guids.Value.ProjectGuid;
                            Guid linkModelGuid = guids.Value.ModelGuid;

                            // If APS didn't return a project GUID, use the host document's
                            if (linkProjectGuid == Guid.Empty)
                                linkProjectGuid = revitProjectId;

                            // Create cloud model path
                            ModelPath linkPath = ModelPathUtils.ConvertCloudGUIDsToCloudPath(
                                "US", linkProjectGuid, linkModelGuid);

                            // Create the link in Revit
                            using (Transaction tx = new Transaction(doc, "Link " + file.Name))
                            {
                                tx.Start();
                                RevitLinkOptions opts = new RevitLinkOptions(false);
                                LinkLoadResult loadResult = RevitLinkType.Create(doc, linkPath, opts);

                                ElementId linkTypeId = loadResult.ElementId;
                                RevitLinkInstance.Create(doc, linkTypeId);
                                tx.Commit();
                            }

                            results.AppendLine($"OK: {file.Name}");
                            successCount++;
                        }
                        catch (Exception ex)
                        {
                            results.AppendLine($"FAILED: {file.Name} — {ex.Message}");
                            failCount++;
                        }
                    }

                    // Step 11: Show results
                    string logPath = System.IO.Path.Combine(
                        System.IO.Path.GetTempPath(),
                        $"HMV_LinkResults_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                    System.IO.File.WriteAllText(logPath, results.ToString());

                    TaskDialog.Show("HMV Tools - Link Results",
                        $"Linked: {successCount}  |  Failed: {failCount}\n\n" +
                        results.ToString() +
                        $"\nFull log: {logPath}");

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