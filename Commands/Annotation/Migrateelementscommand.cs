using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

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

            // ═══════════════════════════════════════════════════
            //  1. GATHER SELECTION
            // ═══════════════════════════════════════════════════

            ICollection<ElementId> selectedIds = uiDoc.Selection
                .GetElementIds();

            if (selectedIds.Count == 0)
            {
                // Prompt user to select
                try
                {
                    var refs = uiDoc.Selection.PickObjects(
                        ObjectType.Element,
                        "Select elements to migrate, then press Finish.");
                    selectedIds = refs.Select(r => r.ElementId).ToList();
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }
            }

            if (selectedIds.Count == 0)
            {
                TaskDialog.Show("HMV Tools", "No elements selected.");
                return Result.Cancelled;
            }

            // ═══════════════════════════════════════════════════
            //  2. BUILD WINDOW DATA (plain classes, no Revit refs)
            // ═══════════════════════════════════════════════════

            // 2a. Open documents (exclude active)
            var openDocs = new List<OpenDocEntry>();
            var docList = new List<Document>();

            int idx = 0;
            foreach (Document doc in uiApp.Application.Documents)
            {
                if (doc.IsLinked) { idx++; continue; }
                if (doc.PathName == srcDoc.PathName) { idx++; continue; }
                if (doc.IsFamilyDocument) { idx++; continue; }

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

            // 2b. Collect transferable views from source
            var viewEntries = CollectTransferableViews(srcDoc);

            // ═══════════════════════════════════════════════════
            //  3. SHOW WINDOW
            // ═══════════════════════════════════════════════════

            var win = new MigrateElementsWindow(
                srcDoc.Title,
                selectedIds.Count,
                openDocs,
                viewEntries);

            bool? dialogResult = win.ShowDialog();
            if (dialogResult != true || win.Settings == null)
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
                TaskDialog.Show("HMV Tools", "Could not resolve target document.");
                return Result.Failed;
            }

            // ═══════════════════════════════════════════════════
            //  4. COMPUTE SHARED COORDINATE TRANSFORM
            // ═══════════════════════════════════════════════════

            Transform coordTransform =
                TransferManager.ComputeSharedCoordinateTransform(srcDoc, tgtDoc);

            string warning;
            if (!TransferManager.ValidateTransform(coordTransform, out warning)
                && warning != null)
            {
                var td = new TaskDialog("HMV Tools – Coordinate Warning");
                td.MainInstruction = "Shared coordinates issue detected";
                td.MainContent = warning
                    + "\n\nDo you want to proceed anyway?";
                td.CommonButtons = TaskDialogCommonButtons.Yes
                                 | TaskDialogCommonButtons.No;
                if (td.Show() != TaskDialogResult.Yes)
                    return Result.Cancelled;
            }
            else if (warning != null)
            {
                // Non-fatal warning (e.g. large offset)
                TaskDialog.Show("HMV Tools – Warning", warning);
            }

            // ═══════════════════════════════════════════════════
            //  5. EXECUTE MIGRATION IN TRANSACTION GROUP
            // ═══════════════════════════════════════════════════

            var result = new TransferResult();

            using (var tg = new TransactionGroup(tgtDoc, "Migrate Elements"))
            {
                tg.Start();

                try
                {
                    // ── 5a. Copy 3D model elements ─────────────
                    TransferManager.CopyModelElements(
                        srcDoc, tgtDoc, selectedIds,
                        coordTransform, result);

                    // ── 5b. Recreate selected views ────────────
                    var viewMap = new Dictionary<ElementId, View>();

                    foreach (int viewIdInt in settings.SelectedViewIds)
                    {
                        var srcViewId = new ElementId(viewIdInt);
                        View srcView = srcDoc.GetElement(srcViewId) as View;
                        if (srcView == null)
                        {
                            result.Warnings.Add(
                                $"Source view Id {viewIdInt} not found, skipped.");
                            continue;
                        }

                        View newView = TransferManager.RecreateView(
                            srcDoc, tgtDoc, srcView,
                            coordTransform, result);

                        if (newView != null)
                            viewMap[srcViewId] = newView;
                    }

                    // ── 5c. Copy annotations (2-step) ──────────
                    if (settings.IncludeAnnotations)
                    {
                        foreach (var kvp in viewMap)
                        {
                            View srcView = srcDoc.GetElement(kvp.Key) as View;
                            View tgtView = kvp.Value;

                            if (srcView == null || tgtView == null) continue;

                            // For plans/sections, annotations may need
                            // the coordinate transform; for drafting/legend
                            // they were already copied in RecreateView.
                            bool is2D =
                                srcView.ViewType == ViewType.DraftingView
                                || srcView.ViewType == ViewType.Legend;

                            if (!is2D)
                            {
                                TransferManager.CopyViewAnnotations(
                                    srcDoc, tgtDoc,
                                    srcView, tgtView,
                                    Transform.Identity,
                                    result);
                            }
                        }
                    }

                    // ── 5d. Reference markers ──────────────────
                    if (settings.IncludeRefMarkers)
                    {
                        foreach (var kvp in viewMap)
                        {
                            View srcView = srcDoc.GetElement(kvp.Key) as View;
                            if (srcView == null) continue;

                            var markers = TransferManager
                                .IdentifyReferenceMarkers(srcDoc, srcView);

                            foreach (var m in markers)
                            {
                                result.Warnings.Add(
                                    $"Ref marker noted (manual recreation may be needed): {m}");
                                result.RefMarkersCreated++;
                            }
                        }
                    }

                    // All succeeded — commit the group
                    tg.Assimilate();
                }
                catch (Exception ex)
                {
                    tg.RollBack();
                    result.Errors.Add($"Fatal error — all changes rolled back: {ex.Message}");
                }
            }

            // ═══════════════════════════════════════════════════
            //  6. SHOW REPORT
            // ═══════════════════════════════════════════════════

            var report = new TaskDialog("HMV Tools – Migration Complete");
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
        //  HELPERS
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Collects all Floor Plans, Sections, Drafting Views, and
        /// Legends from the source document into plain ViewEntry
        /// objects for the UI.
        /// </summary>
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
                        category = "Sections"; break;
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
    }
}