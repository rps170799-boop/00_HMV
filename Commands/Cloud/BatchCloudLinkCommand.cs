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
                    // ─── PHASE 4: Discovery via folder browser ───────────────
                    var hub = System.Threading.Tasks.Task.Run(async () =>
                    {
                        return await aps.GetHubAsync();
                    }).GetAwaiter().GetResult();

                    var browser = new FolderBrowserWindow(aps, hub.Id, hub.Name);
                    browser.ShowDialog();

                    if (browser.DialogResult != true ||
                        string.IsNullOrEmpty(browser.SelectedProjectId) ||
                        string.IsNullOrEmpty(browser.SelectedFolderId))
                    {
                        return Result.Cancelled;
                    }

                    string selectedProjectId = browser.SelectedProjectId;
                    string selectedProjectName = browser.SelectedProjectName;
                    string selectedFolderId = browser.SelectedFolderId;
                    string selectedFolderName = browser.SelectedFolderDisplayPath;

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

                    // Exclude the currently open model from results

                    string currentFileName = doc.Title;
                    if(!currentFileName.EndsWith(".rvt",StringComparison.OrdinalIgnoreCase))
                        currentFileName += ".rvt";
                    rvtFiles.RemoveAll(f => string.Equals(f.Name, currentFileName, StringComparison.OrdinalIgnoreCase));


                    // Ensure all files have default status
                    foreach (var f in rvtFiles) { if (string.IsNullOrEmpty(f.Status)) f.Status = "Not Linked"; }

                    // Step 8: Detect existing links and their status
                    var existingLinks = new Dictionary<string, (int ElementId, string Status)>(
                        StringComparer.OrdinalIgnoreCase);

                    var linkCollector = new FilteredElementCollector(doc);
                    foreach (RevitLinkType linkType in linkCollector.OfClass(typeof(RevitLinkType)).Cast<RevitLinkType>())
                    {
                        string linkName = linkType.Name;
                        string status;
                        try
                        {
                            status = RevitLinkType.IsLoaded(doc, linkType.Id) ? "Loaded" : "Unloaded";
                        }
                        catch
                        {
                            status = "Unloaded";
                        }

                        int eid = linkType.Id.IntegerValue;

                        if (!existingLinks.ContainsKey(linkName))
                            existingLinks[linkName] = (eid, status);

                        string noExt = System.IO.Path.GetFileNameWithoutExtension(linkName);
                        if (!existingLinks.ContainsKey(noExt))
                            existingLinks[noExt] = (eid, status);
                    }

                    // Annotate scan results with link status
                    foreach (var f in rvtFiles)
                    {
                        string nameNoExt = System.IO.Path.GetFileNameWithoutExtension(f.Name);
                        if (existingLinks.TryGetValue(f.Name, out var linkInfo) ||
                            existingLinks.TryGetValue(nameNoExt, out linkInfo))
                        {
                            f.Status = linkInfo.Status;
                            f.LinkTypeElementId = linkInfo.ElementId;
                        }
                    }

                    // Step 9: Show WPF window (all files, with status)
                    var window = new BatchCloudLinkWindow(rvtFiles, selectedProjectName, selectedFolderName);


                    window.ShowDialog();

                    if (window.DialogResult != true)
                        return Result.Cancelled;

                    // ─── Handle RELOAD action ─────────────────────────────────
                    if (window.IsReloadAction && window.ReloadFiles != null && window.ReloadFiles.Count > 0)
                    {
                        var reloadResults = new System.Text.StringBuilder();
                        int reloadOk = 0;
                        int reloadFail = 0;

                        foreach (var f in window.ReloadFiles)
                        {
                            try
                            {
                                ElementId eid = new ElementId(f.LinkTypeElementId);
                                RevitLinkType rlt = doc.GetElement(eid) as RevitLinkType;
                                if (rlt == null)
                                {
                                    reloadResults.AppendLine($"FAILED: {f.Name} — Element not found");
                                    reloadFail++;
                                    continue;
                                }

                                rlt.Reload();

                                reloadResults.AppendLine($"OK: {f.Name}");
                                reloadOk++;
                            }
                            catch (Exception ex)
                            {
                                reloadResults.AppendLine($"FAILED: {f.Name} — {ex.Message}");
                                reloadFail++;
                            }
                        }

                        TaskDialog.Show("HMV Tools - Reload Results",
                            $"Reloaded: {reloadOk}  |  Failed: {reloadFail}\n\n" +
                            reloadResults.ToString());

                        return Result.Succeeded;
                    }

                    // ─── Handle LINK action ───────────────────────────────────
                    if (window.SelectedFiles == null || window.SelectedFiles.Count == 0)
                        return Result.Cancelled;

                    // Step 10: Resolve ALL cloud GUIDs in parallel (fast)
                    var resolvedFiles = System.Threading.Tasks.Task.Run(async () =>
                    {
                        var tasks = window.SelectedFiles.Select(async file =>
                        {
                            var guids = await aps.ResolveCloudGuidsAsync(selectedProjectId, file.Urn);
                            return (File: file, Guids: guids);
                        });
                        return await System.Threading.Tasks.Task.WhenAll(tasks);
                    }).GetAwaiter().GetResult();

                    // Step 11: Create all links in a single transaction (fast)
                    var results = new System.Text.StringBuilder();
                    int successCount = 0;
                    int failCount = 0;

                    int total = resolvedFiles.Length;
                    int current = 0;

                    foreach (var item in resolvedFiles)
                    {
                        current++;
                        try
                        {
                            if (item.Guids == null)
                            {
                                results.AppendLine($"FAILED: {item.File.Name} — Could not resolve cloud GUIDs");
                                failCount++;
                                continue;
                            }

                            Guid linkProjectGuid = item.Guids.Value.ProjectGuid;
                            Guid linkModelGuid = item.Guids.Value.ModelGuid;

                            if (linkProjectGuid == Guid.Empty)
                            {
                                // Try to extract project GUID from the APS project ID (strip "b." prefix)
                                string bareId = selectedProjectId.StartsWith("b.")
                                    ? selectedProjectId.Substring(2)
                                    : selectedProjectId;
                                if (Guid.TryParse(bareId, out Guid parsedGuid))
                                    linkProjectGuid = parsedGuid;
                                else
                                    linkProjectGuid = revitProjectId;
                            }

                            ModelPath linkPath = ModelPathUtils.ConvertCloudGUIDsToCloudPath(
                                "US", linkProjectGuid, linkModelGuid);

                            using (Transaction tx = new Transaction(doc, $"Link {current}/{total}: {item.File.Name}"))
                            {
                                tx.Start();
                                RevitLinkOptions opts = new RevitLinkOptions(false);
                                LinkLoadResult loadResult = RevitLinkType.Create(doc, linkPath, opts);
                                RevitLinkInstance.Create(doc, loadResult.ElementId);
                                tx.Commit();
                            }

                            results.AppendLine($"OK ({current}/{total}): {item.File.Name}");
                            successCount++;
                        }
                        catch (Exception ex)
                        {
                            results.AppendLine($"FAILED ({current}/{total}): {item.File.Name} — {ex.Message}");
                            failCount++;
                        }
                    }



                    // Step 12: Show results
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