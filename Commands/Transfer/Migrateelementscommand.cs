using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class MigrateElementsCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document srcDoc = uiDoc.Document;
            // Allow running with no elements (views-only migration)
            ICollection<ElementId> selectedIds =
                uiDoc.Selection.GetElementIds();
            // ═══════════════════════════════════════════════════
            //  2. BUILD WINDOW DATA
            // ═══════════════════════════════════════════════════

            var openDocs = new List<OpenDocEntry>();
            var docList = new List<Document>();

            int idx = 0;
            foreach (Document doc in uiApp.Application.Documents)
            {
                if (doc.IsLinked || doc.IsFamilyDocument
                    || doc.PathName == srcDoc.PathName)
                { idx++; continue; }

                openDocs.Add(new OpenDocEntry
                {
                    Title = doc.Title,
                    PathName = doc.PathName,
                    Index = idx
                });
                docList.Add(doc);
                idx++;
            }

            if (openDocs.Count == 0)
            {
                TaskDialog.Show("HMV Tools",
                    "No other project documents are open.\n"
                    + "Open the target document first, then re-run.");
                return Result.Cancelled;
            }

            var viewEntries = CollectTransferableViews(srcDoc);
            var sheetEntries = CollectTransferableSheets(srcDoc);

            // ═══════════════════════════════════════════════════
            //  3. SHOW WINDOW
            // ═══════════════════════════════════════════════════

            var win = new MigrateElementsWindow(
                srcDoc.Title,
                selectedIds != null ? selectedIds.Count : 0,
                openDocs,
                viewEntries,
                sheetEntries);

            if (win.ShowDialog() != true || win.Settings == null)
                return Result.Cancelled;

            MigrationSettings settings = win.Settings;

            // Resolve target document
            Document tgtDoc = null;
            int tgtIdx = openDocs[settings.TargetDocIndex].Index;
            int i = 0;
            foreach (Document doc in uiApp.Application.Documents)
            {
                if (i == tgtIdx) { tgtDoc = doc; break; }
                i++;
            }

            if (tgtDoc == null)
            {
                TaskDialog.Show("HMV Tools",
                    "Could not resolve target document.");
                return Result.Failed;
            }

            // ═══════════════════════════════════════════════════
            //  4. SHARED COORDINATE TRANSFORM
            // ═══════════════════════════════════════════════════

            Transform coordTransform =
                SheetHelpers.ComputeSharedCoordinateTransform(
                    srcDoc, tgtDoc);

            string warning;
            if (!SheetHelpers.ValidateTransform(
                    coordTransform, out warning))
            {
                var td = new TaskDialog(
                    "HMV Tools – Coordinate Warning");
                td.MainInstruction =
                    "Shared coordinates issue detected";
                td.MainContent = warning
                    + "\n\nIf the models share the same survey point "
                    + "and are co-located, choose 'Use Identity' to "
                    + "copy without any coordinate offset.\n\n"
                    + "Choose 'Cancel' to abort.";
                td.AddCommandLink(
                    TaskDialogCommandLinkId.CommandLink1,
                    "Use Identity Transform",
                    "Copy elements with no coordinate offset (models are co-located).");
                td.CommonButtons = TaskDialogCommonButtons.Cancel;
                TaskDialogResult tdResult = td.Show();
                if (tdResult == TaskDialogResult.CommandLink1)
                    coordTransform = Transform.Identity;
                else
                    return Result.Cancelled;
            }

            // ═══════════════════════════════════════════════════
            //  5. EXECUTE MIGRATION (TransactionGroup)
            // ═══════════════════════════════════════════════════

            var result = new TransferResult();

            using (var tg = new TransactionGroup(
                tgtDoc, "HMV – Migrate Elements"))
            {
                tg.Start();

                try
                {
                    // ── 5a. Copy 3D model elements ─────────────
                    ICollection<ElementId> copiedModelIds =
                        TransferManager.CopyModelElements(
                            srcDoc, tgtDoc, selectedIds,
                            coordTransform, result);

                    // ── 5b. Copy views (doc-to-doc) ────────────
                    var viewIdsToCopy = settings.SelectedViewIds
                        .Select(id => new ElementId(id))
                        .ToList();

                    Dictionary<ElementId, ElementId> viewMap =
                        TransferManager.CopyViews(
                            srcDoc, tgtDoc,
                            viewIdsToCopy,
                            coordTransform, result,
                            settings.TransferMode);

                    // ── 5c. Copy & assign view templates ───────
                    if (viewMap.Count > 0)
                    {
                        ViewTransferHelper.CopyAndAssignViewTemplates(
                            srcDoc, tgtDoc, viewMap, result);
                    }
                    // ── 5d. Copy view annotations (2nd pass) ───
                    if (viewMap.Count > 0 && settings.IncludeAnnotations)
                    {
                        ViewTransferHelper.CopyViewAnnotations(
                            srcDoc, tgtDoc, viewMap, result);
                    }
                    // ── 5e. Copy category graphic overrides ────
                    if (viewMap.Count > 0)
                    {
                        ViewTransferHelper.CopyCategoryOverrides(
                            srcDoc, tgtDoc, viewMap, result);
                    }

                    // ── 5f. Copy sheets ────────────────────────
                    var sheetIdsToCopy = settings.SelectedSheetIds
                        .Select(id => new ElementId(id))
                        .ToList();

                    if (sheetIdsToCopy.Count > 0)
                    {
                        Dictionary<ElementId, ElementId> sheetMap =
                            TransferManager.CopySheets(
                                srcDoc, tgtDoc,
                                sheetIdsToCopy,
                                coordTransform,
                                result,
                                settings);

                        // Templates and overrides for views
                        // created during sheet viewport copy
                        // are handled inside CopySheets already.
                    }

                    // ── 5g. Reference markers (informational) ──
                    if (settings.IncludeRefMarkers)
                    {
                        foreach (var kvp in viewMap)
                        {
                            View srcView = srcDoc.GetElement(
                                kvp.Key) as View;
                            if (srcView == null) continue;

                            var markers = ViewTransferHelper
                                .IdentifyReferenceMarkers(
                                    srcDoc, srcView);

                            foreach (var m in markers)
                            {
                                result.Warnings.Add(
                                    "Ref marker (manual recreation "
                                    + $"may be needed): {m}");
                                result.RefMarkersNoted++;
                            }
                        }
                    }
                    // ── Debug: check if views are actually new ──
                    foreach (var kvp in viewMap)
                    {
                        View srcV = srcDoc.GetElement(kvp.Key) as View;
                        View tgtV = tgtDoc.GetElement(kvp.Value) as View;
                        if (srcV != null && tgtV != null)
                        {
                            result.Warnings.Add(
                                $"View map: '{srcV.Name}' (src {kvp.Key.IntegerValue}) "
                                + $"→ '{tgtV.Name}' (tgt {kvp.Value.IntegerValue})");
                        }
                    }
                    tg.Assimilate();
                }
                catch (Exception ex)
                {
                    tg.RollBack();
                    result.Errors.Add(
                        "Fatal — all changes rolled back: "
                        + ex.Message);
                }
            }

            // ═══════════════════════════════════════════════════
            //  6. REPORT
            // ═══════════════════════════════════════════════════

            var report = new TaskDialog(
                "HMV Tools – Migration Complete");
            report.MainInstruction = result.Errors.Count == 0
                ? "Migration completed successfully"
                : "Migration completed with errors";
            report.MainContent = result.BuildReport();
            report.CommonButtons = TaskDialogCommonButtons.Ok;
            report.Show();

            return result.Errors.Count == 0
                ? Result.Succeeded
                : Result.Failed;
        }

        // ═══════════════════════════════════════════════════════
        //  COLLECT TRANSFERABLE VIEWS (now includes Legends)
        // ═══════════════════════════════════════════════════════

        private List<ViewEntry> CollectTransferableViews(Document doc)
        {
            var entries = new List<ViewEntry>();

            var views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate
                    && v.ViewType != ViewType.Undefined
                    && v.ViewType != ViewType.Internal
                    && v.ViewType != ViewType.ProjectBrowser
                    && v.ViewType != ViewType.SystemBrowser
                    && v.ViewType != ViewType.Schedule
                    && v.ViewType != ViewType.DrawingSheet
                    && v.ViewType != ViewType.ThreeD)
                .OrderBy(v => v.ViewType.ToString())
                .ThenBy(v => v.Name);

            foreach (View v in views)
            {
                string category;
                switch (v.ViewType)
                {
                    case ViewType.FloorPlan:
                        category = "Floor Plans"; break;
                    case ViewType.CeilingPlan:
                        category = "Ceiling Plans"; break;
                    case ViewType.Section:
                    case ViewType.Detail:
                    case ViewType.Elevation:
                        category = "Sections"; break;
                    case ViewType.DraftingView:
                        category = "Drafting Views"; break;
                    case ViewType.Legend:
                        category = "Legends"; break;
                    default:
                        continue;
                }

                entries.Add(new ViewEntry
                {
                    Id = v.Id.IntegerValue,
                    Name = v.Name,
                    Category = category
                });
            }

            return entries;
        }

        // ═══════════════════════════════════════════════════════
        //  COLLECT TRANSFERABLE SHEETS
        // ═══════════════════════════════════════════════════════

        private List<SheetEntry> CollectTransferableSheets(
            Document doc)
        {
            var entries = new List<SheetEntry>();

            var sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .OrderBy(s => s.SheetNumber);

            foreach (ViewSheet s in sheets)
            {
                entries.Add(new SheetEntry
                {
                    Id = s.Id.IntegerValue,
                    SheetNumber = s.SheetNumber,
                    SheetName = s.Name,
                    ViewportCount = s.GetAllViewports().Count,
                    IsPlaceholder = s.IsPlaceholder
                });
            }

            return entries;
        }
    }
}