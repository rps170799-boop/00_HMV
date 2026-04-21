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
                    if (!currentFileName.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
                        currentFileName += ".rvt";
                    rvtFiles.RemoveAll(f => string.Equals(f.Name, currentFileName, StringComparison.OrdinalIgnoreCase));


                    // Ensure all files have default status
                    foreach (var f in rvtFiles) { if (string.IsNullOrEmpty(f.Status)) f.Status = "Not Linked"; }

                    // Step 8: Detect existing links, their status, AND instance count
                    var existingLinks = new Dictionary<string, (int ElementId, string Status, int Count)>(
                        StringComparer.OrdinalIgnoreCase);

                    // 8A. First, count all physical link instances currently placed in the model
                    var instanceCounts = new Dictionary<ElementId, int>();
                    var linkInstances = new FilteredElementCollector(doc)
                        .OfClass(typeof(RevitLinkInstance))
                        .Cast<RevitLinkInstance>();

                    foreach (var inst in linkInstances)
                    {
                        ElementId typeId = inst.GetTypeId(); // Get the ID of the linked file it belongs to
                        if (instanceCounts.ContainsKey(typeId))
                            instanceCounts[typeId]++;
                        else
                            instanceCounts[typeId] = 1;
                    }

                    // 8B. Scan the loaded Link Types
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

                        // Look up the count we found in Step 8A (default to 0 if none exist)
                        int count = instanceCounts.ContainsKey(linkType.Id) ? instanceCounts[linkType.Id] : 0;

                        if (!existingLinks.ContainsKey(linkName))
                            existingLinks[linkName] = (eid, status, count);

                        string noExt = System.IO.Path.GetFileNameWithoutExtension(linkName);
                        if (!existingLinks.ContainsKey(noExt))
                            existingLinks[noExt] = (eid, status, count);
                    }

                    // Annotate scan results with link status AND count
                    foreach (var f in rvtFiles)
                    {
                        string nameNoExt = System.IO.Path.GetFileNameWithoutExtension(f.Name);
                        if (existingLinks.TryGetValue(f.Name, out var linkInfo) ||
                            existingLinks.TryGetValue(nameNoExt, out linkInfo))
                        {
                            f.Status = linkInfo.Status;
                            f.LinkTypeElementId = linkInfo.ElementId;
                            f.InstanceCount = linkInfo.Count; // Set the count!
                        }
                        else
                        {
                            f.InstanceCount = 0;
                        }
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

                    // Step 9: Show WPF window — loop until user clicks Cancel
                    

                    while (true)
                    {
                        var window = new BatchCloudLinkWindow(rvtFiles, selectedProjectName, selectedFolderName);
                        window.ShowDialog();

                        if (window.DialogResult != true)
                            break;

                        // ─── Handle RELOAD action ─────────────────────────────
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
                                    window.UpdateRowStatus(f.Name, "Loaded");

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

                            window.RefreshAfterOperation();
                            continue;
                        }

                        // ─── Handle LINK action ───────────────────────────────
                        if (window.SelectedFiles != null && window.SelectedFiles.Count > 0)
                        {
                            var resolvedFiles2 = System.Threading.Tasks.Task.Run(async () =>
                            {
                                var tasks = window.SelectedFiles.Select(async file =>
                                {
                                    var guids = await aps.ResolveCloudGuidsAsync(selectedProjectId, file.Urn);
                                    return (File: file, Guids: guids);
                                });
                                return await System.Threading.Tasks.Task.WhenAll(tasks);
                            }).GetAwaiter().GetResult();

                            var linkResults = new System.Text.StringBuilder();
                            int successCount = 0;
                            int failCount = 0;
                            int total = resolvedFiles2.Length;
                            int current = 0;

                            foreach (var item in resolvedFiles2)
                            {
                                current++;
                                try
                                {
                                    if (item.Guids == null)
                                    {
                                        linkResults.AppendLine($"FAILED: {item.File.Name} — Could not resolve cloud GUIDs");
                                        failCount++;
                                        continue;
                                    }

                                    Guid linkProjectGuid = item.Guids.Value.ProjectGuid;
                                    Guid linkModelGuid = item.Guids.Value.ModelGuid;

                                    if (linkProjectGuid == Guid.Empty)
                                    {
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

                                        // 1. Setup Defaults
                                        ImportPlacement placementEnum = ImportPlacement.Shared;
                                        bool needsPbpWorkaround = false;

                                        // 2. Map the UI selection
                                        if (item.File.Placement == "Origin to Origin")
                                        {
                                            placementEnum = ImportPlacement.Origin;
                                        }
                                        else if (item.File.Placement == "Project Base Point")
                                        {
                                            // Fake it! Tell the API to use Origin so it doesn't crash...
                                            placementEnum = ImportPlacement.Origin;
                                            // ...but set our flag so we remember to slide it over!
                                            needsPbpWorkaround = true;
                                        }

                                        // 3. Create the instance
                                        RevitLinkInstance instance = RevitLinkInstance.Create(doc, loadResult.ElementId, placementEnum);

                                        // 4. Apply the Workaround if requested
                                        if (needsPbpWorkaround)
                                        {
                                            AlignToProjectBasePoint(doc, instance);
                                        }

                                        tx.Commit();
                                    }

                                    window.UpdateRowStatus(item.File.Name, "Loaded");

                                    linkResults.AppendLine($"OK ({current}/{total}): {item.File.Name}");
                                    successCount++;
                                }
                                catch (Exception ex)
                                {
                                    linkResults.AppendLine($"FAILED ({current}/{total}): {item.File.Name} — {ex.Message}");
                                    failCount++;
                                }
                            }

                            string logPath = System.IO.Path.Combine(
                                System.IO.Path.GetTempPath(),
                                $"HMV_LinkResults_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                            System.IO.File.WriteAllText(logPath, linkResults.ToString());

                            TaskDialog.Show("HMV Tools - Link Results",
                                $"Linked: {successCount}  |  Failed: {failCount}\n\n" +
                                linkResults.ToString() +
                                $"\nFull log: {logPath}");
                            // Silently set shared sites for newly linked files
                            try
                            {
                                var linkedNames = resolvedFiles2
                                    .Where(r => r.Guids != null)
                                    .Select(r => r.File.Name).ToList();

                                using (Transaction siteTx = new Transaction(doc, "Set Shared Sites"))
                                {
                                    siteTx.Start();
                                    siteTx.Commit();
                                }
                            }
                            catch { }

                            RefreshUIStatusOnly(doc, rvtFiles, window.FileRows);

                            window.RefreshAfterOperation();
                            continue;
                        }

                        break;
                    }
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


        private void RefreshUIStatusOnly(Document doc, List<CloudRevitFile> rvtFiles, IEnumerable<RvtFileRow> uiRows)
        {
            // 1. Count all physical link instances
            var instanceCounts = new Dictionary<ElementId, int>();
            var linkInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>();

            foreach (var inst in linkInstances)
            {
                ElementId typeId = inst.GetTypeId();
                if (instanceCounts.ContainsKey(typeId))
                    instanceCounts[typeId]++;
                else
                    instanceCounts[typeId] = 1;
            }

            // 2. Get all loaded Link Types
            var linkTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkType))
                .Cast<RevitLinkType>()
                .ToList();

            var typeMap = new Dictionary<string, (int Id, string Status, int Count)>(StringComparer.OrdinalIgnoreCase);
            foreach (var lt in linkTypes)
            {
                string status = "Unloaded";
                try { status = RevitLinkType.IsLoaded(doc, lt.Id) ? "Loaded" : "Unloaded"; } catch { }
                int count = instanceCounts.ContainsKey(lt.Id) ? instanceCounts[lt.Id] : 0;

                typeMap[lt.Name] = (lt.Id.IntegerValue, status, count);
                typeMap[System.IO.Path.GetFileNameWithoutExtension(lt.Name)] = (lt.Id.IntegerValue, status, count);
            }

            // 3. UPDATE THE DATABASE (rvtFiles) so new windows get the correct data
            foreach (var f in rvtFiles)
            {
                string rowNameNoExt = System.IO.Path.GetFileNameWithoutExtension(f.Name);
                var matchingType = linkTypes.FirstOrDefault(lt =>
                    lt.Name.Equals(f.Name, StringComparison.OrdinalIgnoreCase) ||
                    lt.Name.IndexOf(rowNameNoExt, StringComparison.OrdinalIgnoreCase) >= 0);

                if (matchingType != null)
                {
                    f.Status = typeMap.ContainsKey(matchingType.Name) ? typeMap[matchingType.Name].Status : "Loaded";
                    f.LinkTypeElementId = matchingType.Id.IntegerValue;
                    f.InstanceCount = instanceCounts.ContainsKey(matchingType.Id) ? instanceCounts[matchingType.Id] : 0;
                }
            }

            // 4. Update the visual rows (just in case)
            if (uiRows != null)
            {
                foreach (var row in uiRows)
                {
                    string rowNameNoExt = System.IO.Path.GetFileNameWithoutExtension(row.Name);
                    var matchingType = linkTypes.FirstOrDefault(lt =>
                        lt.Name.Equals(row.Name, StringComparison.OrdinalIgnoreCase) ||
                        lt.Name.IndexOf(rowNameNoExt, StringComparison.OrdinalIgnoreCase) >= 0);

                    if (matchingType != null)
                    {
                        row.Status = typeMap.ContainsKey(matchingType.Name) ? typeMap[matchingType.Name].Status : "Loaded";
                        row.LinkTypeElementId = matchingType.Id.IntegerValue;
                        row.InstanceCount = instanceCounts.ContainsKey(matchingType.Id) ? instanceCounts[matchingType.Id] : 0;
                    }
                }
            }
        }
        /// <summary>
        /// Moves a linked instance so its Project Base Point matches the host's Project Base Point.
        /// Assumes the link was originally placed Origin-to-Origin.
        /// </summary>
        private void AlignToProjectBasePoint(Document hostDoc, RevitLinkInstance linkInstance)
        {
            Document linkDoc = linkInstance.GetLinkDocument();
            if (linkDoc == null) return;

            XYZ hostPbp = GetPbpInternalPosition(hostDoc);
            XYZ linkPbp = GetPbpInternalPosition(linkDoc);

            // Calculate the vector needed to slide the link into place
            XYZ translationVector = hostPbp - linkPbp;

            // If it's not already perfectly aligned, move the instance!
            if (!translationVector.IsAlmostEqualTo(XYZ.Zero))
            {
                ElementTransformUtils.MoveElement(hostDoc, linkInstance.Id, translationVector);
            }
        }

        /// <summary>
        /// Safely extracts the XYZ coordinates of a Project Base Point relative to the Internal Origin.
        /// </summary>
        private XYZ GetPbpInternalPosition(Document doc)
        {
            Element pbp = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                .WhereElementIsNotElementType()
                .FirstElement();

            if (pbp != null)
            {
                // A BoundingBox correctly captures the precise XYZ location relative to the Internal Origin
                BoundingBoxXYZ bbox = pbp.get_BoundingBox(null);
                if (bbox != null) return bbox.Min;
            }

            // Fallback if the model is somehow missing a PBP (extremely rare)
            return XYZ.Zero;
        }

    }
}