using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;

namespace HMVTools
{
    // ── TransferManager ────────────────────────────────────────

    public static class TransferManager
    {
        // ═══════════════════════════════════════════════════════
        //  2. COPY 3D MODEL ELEMENTS
        // ═══════════════════════════════════════════════════════

        public static ICollection<ElementId> CopyModelElements(
            Document source,
            Document target,
            ICollection<ElementId> elementIds,
            Transform coordTransform,
            TransferResult result,
            string groupName = "HMV – Imported Model Elements")
        {
            var validIds = new List<ElementId>();

            foreach (var id in elementIds)
            {
                Element el = source.GetElement(id);
                if (el == null) continue;

                if (el.OwnerViewId != ElementId.InvalidElementId)
                {
                    result.Warnings.Add(
                        $"Skipped view-specific element: "
                        + $"{el.Name} (Id {id.IntegerValue})");
                    continue;
                }

                if (el.Category != null &&
                    (BuiltInCategory)el.Category.Id.IntegerValue
                        == BuiltInCategory.OST_Viewers)
                    continue;

                validIds.Add(id);
            }

            if (validIds.Count == 0) return new List<ElementId>();

            var opts = new CopyPasteOptions();
            opts.SetDuplicateTypeNamesHandler(
                new UseDestinationTypesHandler());

            using (var t = new Transaction(target,
                "HMV – Copy Model Elements"))
            {
                t.Start();
                try
                {
                    var copied = ElementTransformUtils.CopyElements(
                        source, validIds, target,
                        coordTransform, opts);

                    result.ElementsCopied = copied.Count;

                    // ── Group all copied model elements ───────────────
                    if (copied.Count > 0)
                    {
                        try
                        {
                            Group grp = target.Create.NewGroup(copied);

                            // Rename the group type if possible
                            if (grp != null)
                            {
                                GroupType grpType =
                                    target.GetElement(grp.GetTypeId())
                                        as GroupType;

                                if (grpType != null)
                                {
                                    // Ensure name is unique
                                    string uniqueName = groupName;
                                    var existingNames = new HashSet<string>(
                                        new FilteredElementCollector(target)
                                            .OfClass(typeof(GroupType))
                                            .Cast<GroupType>()
                                            .Select(g => g.Name));

                                    int n = 2;
                                    while (existingNames.Contains(uniqueName))
                                        uniqueName = $"{groupName} ({n++})";

                                    grpType.Name = uniqueName;
                                }

                                result.Warnings.Add(
                                    $"Model elements grouped: "
                                    + $"{copied.Count} elements → "
                                    + $"Group '{grpType?.Name ?? "(unnamed)"}' "
                                    + $"(Id {grp.Id.IntegerValue})");
                            }
                        }
                        catch (Exception gex)
                        {
                            result.Warnings.Add(
                                $"Elements copied but grouping failed: "
                                + gex.Message);
                        }
                    }

                    t.Commit();
                    return copied;
                }
                catch (Exception)
                {
                    t.RollBack();
                    return new List<ElementId>();
                }
            }
        }


        // ═══════════════════════════════════════════════════════
        //  3. COPY VIEWS (DOCUMENT-TO-DOCUMENT)
        //     Now supports Create vs Update mode.
        // ═══════════════════════════════════════════════════════

        public static Dictionary<ElementId, ElementId> CopyViews(
            Document source,
            Document target,
            ICollection<ElementId> viewIds,
            Transform coordTransform,
            TransferResult result,
            ViewTransferMode mode = ViewTransferMode.Create)
        {
            var viewMap = new Dictionary<ElementId, ElementId>();
            if (viewIds == null || viewIds.Count == 0)
                return viewMap;

            var spatialIds = new List<ElementId>();
            var flatIds = new List<ElementId>();

            foreach (var id in viewIds)
            {
                View v = source.GetElement(id) as View;
                if (v == null)
                {
                    result.Warnings.Add(
                        $"View Id {id.IntegerValue} not found, skipped.");
                    continue;
                }

                switch (v.ViewType)
                {
                    case ViewType.FloorPlan:
                    case ViewType.CeilingPlan:
                    case ViewType.Section:
                    case ViewType.Detail:
                    case ViewType.Elevation:
                        spatialIds.Add(id);
                        break;

                    case ViewType.DraftingView:
                    case ViewType.Legend:
                        flatIds.Add(id);
                        break;

                    default:
                        result.Warnings.Add(
                            $"Unsupported view type '{v.ViewType}' "
                            + $"for '{v.Name}', skipped.");
                        break;
                }
            }

            var opts = new CopyPasteOptions();
            opts.SetDuplicateTypeNamesHandler(
                new UseDestinationTypesHandler());

            if (spatialIds.Count > 0)
            {
                ViewFactory.CopyViewBatch(source, target,
                    spatialIds, coordTransform,
                    opts, viewMap, result,
                    "Plan/Section/Elevation", mode);
            }

            if (flatIds.Count > 0)
            {
                var legendIds   = new List<ElementId>();
                var draftingIds = new List<ElementId>();

                foreach (var id in flatIds)
                {
                    View v = source.GetElement(id) as View;
                    if (v?.ViewType == ViewType.DraftingView)
                        draftingIds.Add(id);
                    else
                        legendIds.Add(id);
                }

                // Drafting views: each needs its OWN standalone transaction
                // — MUST NOT run inside CopyViewBatch's outer transaction
                var draftSettings = new MigrationSettings
                {
                    TransferMode = mode,
                    SheetOptions = new SheetCopyOptions()
                };

                foreach (var id in draftingIds)
                {
                    View srcV = source.GetElement(id) as View;
                    if (srcV == null) continue;

                    ElementId newId = ViewFactory.CopyOrUpdateSingleView(
                        source, target, srcV,
                        Transform.Identity, result, draftSettings);

                    if (newId != null && newId != ElementId.InvalidElementId)
                        viewMap[id] = newId;
                }

                // Legends: go through normal batch
                if (legendIds.Count > 0)
                {
                    ViewFactory.CopyViewBatch(source, target,
                        legendIds, Transform.Identity,
                        opts, viewMap, result,
                        "Legend", mode);
                }
            }

            return viewMap;
        }

        // ═══════════════════════════════════════════════════════
        //  7. COPY SHEETS
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Top-level method that copies sheets from source to target.
        /// For each sheet: creates or updates the sheet, then copies
        /// viewports, schedules, title blocks, guide grids, revisions.
        /// Mirrors pyRevit copy_sheet flow.
        /// </summary>
        public static Dictionary<ElementId, ElementId> CopySheets(
            Document source,
            Document target,
            ICollection<ElementId> sheetIds,
            Transform coordTransform,
            TransferResult result,
            MigrationSettings settings)
        {
            var sheetMap = new Dictionary<ElementId, ElementId>();
            // Master view map: source viewId → target viewId
            // Shared across sheets so a view placed on multiple
            // sheets is only copied once.
            var viewMap = new Dictionary<ElementId, ElementId>();

            if (sheetIds == null || sheetIds.Count == 0)
                return sheetMap;

            var opts = settings.SheetOptions;
            result.Warnings.Add(
                $"[CopySheets] called with {sheetIds.Count} sheet(s).");
            foreach (var srcSheetId in sheetIds)
            {
                ViewSheet srcSheet = source.GetElement(srcSheetId)
                    as ViewSheet;
                if (srcSheet == null)
                {
                    result.Warnings.Add(
                        $"Sheet Id {srcSheetId.IntegerValue} "
                        + "not found, skipped.");
                    continue;
                    }
                    Trace.WriteLine(
                        $"[DBG-SH2] Processing sheet '{srcSheet.SheetNumber} - {srcSheet.Name}' "
                        + $"id={srcSheetId.IntegerValue} isPlaceholder={srcSheet.IsPlaceholder}");
                try
                {
                    ViewSheet destSheet = CopySingleSheet(
                        source, target, srcSheet,
                        coordTransform, viewMap,
                        result, settings);

                    if (destSheet != null)
                        sheetMap[srcSheetId] = destSheet.Id;
                }
                catch (Exception ex)
                {
                    result.Errors.Add(
                        $"Sheet '{srcSheet.SheetNumber} - "
                        + $"{srcSheet.Name}' failed: {ex.Message}");
                }
            }

            return sheetMap;
        }

        // ───────────────────────────────────────────────────────────────
        //  8. REPLACE: CopySingleSheet
        //     Search: "private static ViewSheet CopySingleSheet"
        //     Fix: viewport flow now always runs, better error handling
        // ───────────────────────────────────────────────────────────────

        private static ViewSheet CopySingleSheet(
            Document source, Document target,
            ViewSheet srcSheet,
            Transform coordTransform,
            Dictionary<ElementId, ElementId> viewMap,
            TransferResult result,
            MigrationSettings settings)
        {
            var opts = settings.SheetOptions;
            ViewTransferMode mode = settings.TransferMode;

            // ── Check for existing sheet ──────────────────────
            ViewSheet destSheet = ViewFactory.FindMatchingView(
                target, srcSheet) as ViewSheet;
            Trace.WriteLine(
                $"[DBG-SH3] '{srcSheet.SheetNumber}': "
                + $"existingMatch={(destSheet != null ? destSheet.Id.IntegerValue.ToString() : "null")} "
                + $"mode={mode}");
            if (destSheet != null)
            {
                if (mode == ViewTransferMode.Update)
                {
                    using (var t = new Transaction(target,
                        "HMV – Update Sheet Contents"))
                    {
                        t.Start();
                        ViewContentCopier.ClearViewContents(target, destSheet, opts);
                        ViewContentCopier.CopyViewContentsViewToView(
                            source, srcSheet, target, destSheet, opts);
                        t.Commit();
                    }
                    result.SheetsUpdated++;
                }
                else
                {
                    result.Warnings.Add(
                        $"Sheet '{srcSheet.SheetNumber}' already "
                        + "exists in target.");
                }
            }
            else
            {
                // ── Create new sheet ──────────────────────────
                using (var t = new Transaction(target,
                    "HMV – Create Sheet"))
                {
                    t.Start();

                    if (!srcSheet.IsPlaceholder
                        || opts.PlaceholdersAsSheets)
                    {
                        destSheet = ViewSheet.Create(
                            target, ElementId.InvalidElementId);
                    }
                    else
                    {
                        destSheet = ViewSheet.CreatePlaceholder(
                            target);
                    }

                    try
                    {
                        destSheet.Name = srcSheet.Name;
                        destSheet.SheetNumber = srcSheet.SheetNumber;
                    }
                    catch (Exception ex)
                    {
                        result.Warnings.Add(
                            $"Could not set sheet name/number: "
                            + ex.Message);
                    }

                    t.Commit();
                }
                result.SheetsCreated++;
                Trace.WriteLine(
                    $"[DBG-SH4] '{srcSheet.SheetNumber}': created destSheet id={destSheet.Id.IntegerValue}");
                // Copy sheet-level content (title block, schedules)
                // in a SEPARATE transaction after sheet creation
                using (var t = new Transaction(target,
                    "HMV – Copy Sheet Contents"))
                {
                    t.Start();
                    try
                    {
                        ViewContentCopier.CopyViewContentsViewToView(
                            source, srcSheet, target, destSheet, opts);
                        t.Commit();
                        Trace.WriteLine(
                            $"[DBG-SH5] '{srcSheet.SheetNumber}': sheet content copied OK.");
                    }
                    catch (Exception ex)
                    {
                        t.RollBack();
                        Trace.WriteLine(
                            $"[DBG-SH5] '{srcSheet.SheetNumber}': sheet content FAILED: {ex.Message}");
                    }
                }
            }

            if (destSheet == null) return null;

            // ── Viewports ─────────────────────────────────────
            // Always attempt viewport copy for non-placeholder sheets
            Trace.WriteLine(
                $"[DBG-SH6] '{srcSheet.SheetNumber}': "
                + $"isPlaceholder={destSheet.IsPlaceholder} "
                + $"CopyViewports={opts.CopyViewports} "
                + $"srcViewportCount={srcSheet.GetAllViewports().Count}");
            if (!destSheet.IsPlaceholder && opts.CopyViewports)
            {
                try
                {
                    ViewportPlacer.CopySheetViewports(
                        source, srcSheet,
                        target, destSheet,
                        coordTransform, viewMap,
                        result, settings);
                }
                catch (Exception ex)
                {
                    result.Errors.Add(
                        $"Viewport copy failed for sheet "
                        + $"'{srcSheet.SheetNumber}': {ex.Message}");
                }
            }

            // ── Annotations for views created during viewport placement ──
            try
            {
                ViewTransferHelper.CopyViewAnnotations(source, target, viewMap, result);
                Trace.WriteLine(
                    $"[DBG-SH-ANN] Sheet '{srcSheet.SheetNumber}': "
                    + $"CopyViewAnnotations completed, "
                    + $"annotationsCopied={result.AnnotationsCopied}");
            }
            catch (Exception ex)
            {
                result.Warnings.Add(
                    $"[DBG-SH-ANN] Sheet '{srcSheet.SheetNumber}': "
                    + $"CopyViewAnnotations failed: {ex.Message}");
            }

            // ── Revisions ────────────────────────────────────
            if (opts.CopyRevisions)
            {
                SheetHelpers.CopySheetRevisions(
                    source, srcSheet, target, destSheet, result);
            }

            return destSheet;
        }
    }
}
