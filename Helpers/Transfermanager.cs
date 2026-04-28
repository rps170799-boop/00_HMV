using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace HMVTools
{
    // ── Duplicate type handler ─────────────────────────────────

    public class UseDestinationTypesHandler : IDuplicateTypeNamesHandler
    {
        public DuplicateTypeAction OnDuplicateTypeNamesFound(
            DuplicateTypeNamesHandlerArgs args)
        {
            return DuplicateTypeAction.UseDestinationTypes;
        }
    }

    // ── Transfer result container ──────────────────────────────

    public class TransferResult
    {
        public int ElementsCopied { get; set; }
        public int ViewsCreated { get; set; }
        public int ViewsUpdated { get; set; }
        public int SheetsCreated { get; set; }
        public int SheetsUpdated { get; set; }
        public int ViewportsCopied { get; set; }
        public int TemplatesRemapped { get; set; }
        public int RefMarkersNoted { get; set; }
        public int AnnotationsCopied { get; set; }
        public List<string> Warnings { get; set; } = new List<string>();
        public List<string> Errors { get; set; } = new List<string>();

        public string BuildReport()
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("═══  MIGRATION REPORT  ═══");
            sb.AppendLine();
            sb.AppendLine($"  Model elements copied:    {ElementsCopied}");
            sb.AppendLine($"  Views created:            {ViewsCreated}");
            sb.AppendLine($"  Views updated:            {ViewsUpdated}");
            sb.AppendLine($"  Sheets created:           {SheetsCreated}");
            sb.AppendLine($"  Sheets updated:           {SheetsUpdated}");
            sb.AppendLine($"  Viewports placed:         {ViewportsCopied}");
            sb.AppendLine($"  Templates remapped:       {TemplatesRemapped}");
            sb.AppendLine($"  Ref markers identified:   {RefMarkersNoted}");
            sb.AppendLine($"  Annotations copied:       {AnnotationsCopied}");

            if (Warnings.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("── WARNINGS ──");
                foreach (var w in Warnings)
                    sb.AppendLine($"  ⚠  {w}");
            }
            if (Errors.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("── ERRORS ──");
                foreach (var e in Errors)
                    sb.AppendLine($"  ✗  {e}");
            }
            return sb.ToString();
        }
    }

    // ── TransferManager ────────────────────────────────────────

    public static class TransferManager
    {
        // ═══════════════════════════════════════════════════════
        //  1. SHARED COORDINATES TRANSFORM
        // ═══════════════════════════════════════════════════════

        public static Transform ComputeSharedCoordinateTransform(
            Document source, Document target)
        {
            Transform tSrc = source.ActiveProjectLocation
                                   .GetTotalTransform();
            Transform tTgt = target.ActiveProjectLocation
                                   .GetTotalTransform();

            return tTgt.Inverse.Multiply(tSrc);
        }

        public static bool ValidateTransform(
            Transform t, out string warning)
        {
            warning = null;

            if (double.IsNaN(t.Origin.X) ||
                double.IsNaN(t.Origin.Y) ||
                double.IsNaN(t.Origin.Z))
            {
                warning = "Transform contains NaN — shared coordinates "
                        + "may not be configured in one or both files.";
                return false;
            }

            double dist = t.Origin.GetLength();
            if (dist > 32808)
            {
                warning = $"Translation is {dist:F0} ft "
                        + $"(~{dist * 0.3048:F0} m). "
                        + "Shared coordinates may be misconfigured.";
            }

            return true;
        }

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
                CopyViewBatch(source, target,
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

                    ElementId newId = CopyOrUpdateSingleView(
                        source, target, srcV,
                        Transform.Identity, result, draftSettings);

                    if (newId != null && newId != ElementId.InvalidElementId)
                        viewMap[id] = newId;
                }

                // Legends: go through normal batch
                if (legendIds.Count > 0)
                {
                    CopyViewBatch(source, target,
                        legendIds, Transform.Identity,
                        opts, viewMap, result,
                        "Legend", mode);
                }
            }

            return viewMap;
        }

        private static void CopyViewBatch(
            Document source,
            Document target,
            List<ElementId> ids,
            Transform transform,
            CopyPasteOptions opts,
            Dictionary<ElementId, ElementId> viewMap,
            TransferResult result,
            string batchLabel,
            ViewTransferMode mode = ViewTransferMode.Create)
        {
            // ── Separate force-path views (ViewSection) from the rest ──
            // ForceCreateSection/ForceCreatePlan open their OWN
            // Transactions and cannot run inside an outer Transaction.
            var forcePathIds  = new List<ElementId>();
            var regularIds    = new List<ElementId>();

            foreach (var id in ids)
            {
                View v = source.GetElement(id) as View;
                if (v is ViewSection)
                    forcePathIds.Add(id);
                else
                    regularIds.Add(id);
            }

            int batchSuccesses = 0;
            int failures       = 0;

            // ══════════════════════════════════════════════════════
            //  PASS 1 – regular views (inside outer Transaction)
            // ══════════════════════════════════════════════════════
            if (regularIds.Count > 0)
            {
                using (var t = new Transaction(target,
                    $"HMV – Copy {batchLabel} Views"))
                {
                    t.Start();

                    var existingViewIds = new HashSet<int>(
                        new FilteredElementCollector(target)
                            .OfClass(typeof(View))
                            .Select(v => v.Id.IntegerValue));

                    foreach (var srcId in regularIds)
                    {
                        string viewName = "(unknown)";
                        View srcV = null;
                        try
                        {
                            srcV = source.GetElement(srcId) as View;
                            if (srcV != null) viewName = srcV.Name;
                        }
                        catch { }

                        // ── UPDATE MODE ───────────────────────────────
                        if (mode == ViewTransferMode.Update)
                        {
                            View existing = FindMatchingView(target, srcV);
                            if (existing != null)
                            {
                                using (var sub = new SubTransaction(target))
                                {
                                    try
                                    {
                                        sub.Start();
                                        ClearViewContents(target, existing,
                                            new SheetCopyOptions());
                                        CopyViewContentsViewToView(
                                            source, srcV, target, existing,
                                            new SheetCopyOptions());
                                        CopyViewProperties(srcV, existing);
                                        viewMap[srcId] = existing.Id;
                                        result.ViewsUpdated++;
                                        batchSuccesses++;
                                        sub.Commit();
                                    }
                                    catch (Exception ex)
                                    {
                                        try { if (sub.HasStarted() && !sub.HasEnded()) sub.RollBack(); } catch { }
                                        failures++;
                                        result.Warnings.Add(
                                            $"Update failed for '{viewName}': {ex.Message}");
                                    }
                                }
                                continue;
                            }
                            // No match → fall through to create
                        }

                        // ── DRAFTING → standalone path (each needs its own tx) ─
                        if (srcV != null && srcV.ViewType == ViewType.DraftingView)
                        {
                            var draftSettings = new MigrationSettings
                            {
                                TransferMode = mode,
                                SheetOptions = new SheetCopyOptions()
                            };

                            ElementId draftId = CopyOrUpdateSingleView(
                                source, target, srcV,
                                transform, result, draftSettings);

                            if (draftId != null
                                && draftId != ElementId.InvalidElementId)
                            {
                                viewMap[srcId] = draftId;
                                batchSuccesses++;
                            }
                            else
                            {
                                failures++;
                            }
                            continue;
                        }

                        // ── CREATE MODE (normal copy) ─────────────────
                        using (var sub = new SubTransaction(target))
                        {
                            try
                            {
                                sub.Start();

                                var single = new List<ElementId> { srcId };

                                ICollection<ElementId> copied =
                                    ElementTransformUtils.CopyElements(
                                        source, single, target,
                                        transform, opts);

                                if (copied != null && copied.Count > 0)
                                {
                                    ElementId newId = copied.First();

                                    if (existingViewIds.Contains(
                                            newId.IntegerValue))
                                    {
                                        sub.RollBack();
                                        failures++;
                                        result.Warnings.Add(
                                            $"View '{viewName}': "
                                            + "returned existing view.");
                                    }
                                    else
                                    {
                                        viewMap[srcId] = newId;
                                        result.ViewsCreated++;
                                        batchSuccesses++;
                                        sub.Commit();
                                    }
                                }
                                else
                                {
                                    sub.RollBack();
                                    failures++;
                                    result.Warnings.Add(
                                        $"View '{viewName}' produced "
                                        + "no result, skipped.");
                                }
                            }
                            catch (Exception ex)
                            {
                                try
                                {
                                    if (sub.HasStarted() && !sub.HasEnded())
                                        sub.RollBack();
                                }
                                catch { }

                                failures++;
                                result.Warnings.Add(
                                    $"View '{viewName}' failed: {ex.Message}");
                            }
                        }
                    }

                    if (batchSuccesses > 0 || failures < regularIds.Count)
                        t.Commit();
                    else
                    {
                        t.RollBack();
                        result.Errors.Add(
                            $"Copy {batchLabel} views: "
                            + $"all {regularIds.Count} regular views failed.");
                    }
                }
            }

            // ══════════════════════════════════════════════════════
            //  PASS 2 – force-path views (sections/details/elevations)
            //  Each ForceCreateSection opens its OWN transactions,
            //  so they MUST run outside any outer Transaction.
            // ══════════════════════════════════════════════════════
            foreach (var srcId in forcePathIds)
            {
                string viewName = "(unknown)";
                View srcV = null;
                try
                {
                    srcV = source.GetElement(srcId) as View;
                    if (srcV != null) viewName = srcV.Name;
                }
                catch { }

                // UPDATE MODE: find existing and update in-place
                if (mode == ViewTransferMode.Update)
                {
                    View existing = FindMatchingView(target, srcV);
                    if (existing != null)
                    {
                        using (var t = new Transaction(target,
                            $"HMV – Update Section '{viewName}'"))
                        {
                            t.Start();
                            try
                            {
                                ClearViewContents(target, existing,
                                    new SheetCopyOptions());
                                CopyViewContentsViewToView(
                                    source, srcV, target, existing,
                                    new SheetCopyOptions());
                                CopyViewProperties(srcV, existing);
                                viewMap[srcId] = existing.Id;
                                result.ViewsUpdated++;
                                batchSuccesses++;
                                t.Commit();
                            }
                            catch (Exception ex)
                            {
                                t.RollBack();
                                failures++;
                                result.Warnings.Add(
                                    $"Update failed for '{viewName}': {ex.Message}");
                            }
                        }
                        continue;
                    }
                    // No match → fall through to force-create
                }

                // ForceCreateView manages its own Transactions internally
                ElementId forced = ForceCreateView(
                    source, target, srcId, result);

                if (forced != null
                    && forced != ElementId.InvalidElementId)
                {
                    viewMap[srcId] = forced;
                    result.ViewsCreated++;
                    batchSuccesses++;
                }
                else
                {
                    failures++;
                    result.Warnings.Add(
                        $"View '{viewName}': could not duplicate in target.");
                }
            }
        }


        // ═══════════════════════════════════════════════════════
        //  3b. VIEW MATCHING / CONTENT TRANSFER (pyRevit pattern)
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Finds a view in the target document that matches by
        /// ViewType and Name. For sheets, also matches SheetNumber.
        /// Mirrors pyRevit find_matching_view.
        /// </summary>
        public static View FindMatchingView(
            Document target, View sourceView)
        {
            if (sourceView == null) return null;

            foreach (View v in new FilteredElementCollector(target)
                .OfClass(typeof(View))
                .Cast<View>())
            {
                if (v.ViewType != sourceView.ViewType) continue;
                if (v.Name != sourceView.Name) continue;

                if (sourceView.ViewType == ViewType.DrawingSheet)
                {
                    var srcSheet = sourceView as ViewSheet;
                    var tgtSheet = v as ViewSheet;
                    if (srcSheet != null && tgtSheet != null
                        && tgtSheet.SheetNumber == srcSheet.SheetNumber)
                        return v;
                }
                else
                {
                    return v;
                }
            }
            return null;
        }

        /// <summary>
        /// Collects all element IDs owned by a view, filtering out
        /// viewports, extent elements, guide grids, and optionally
        /// title blocks and schedules based on options.
        /// Mirrors pyRevit get_view_contents.
        /// </summary>
        public static List<ElementId> GetViewContents(
            Document doc, View view, SheetCopyOptions options)
        {
            var result = new List<ElementId>();

            var elements = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .ToElements();

            foreach (Element el in elements)
            {
                if (el == null) continue;

                // Skip title blocks if option is off
                if (el.Category != null
                    && el.Category.Name == "Title Blocks"
                    && !options.CopyTitleblock)
                    continue;

                // Skip schedules if option is off
                if (el is ScheduleSheetInstance
                    && !options.CopySchedules)
                    continue;

                // Always skip viewports and extent elements
                if (el is Viewport) continue;
                if (el.Name != null && el.Name.Contains("ExtentElem"))
                    continue;

                // Skip guide grids
                if (el.Category != null
                    && el.Category.Name.IndexOf("guide",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    continue;

                // Skip view references
                if (el.Category != null
                    && string.Equals(el.Category.Name, "views",
                        StringComparison.OrdinalIgnoreCase))
                    continue;

                result.Add(el.Id);
            }
            return result;
        }

        /// <summary>
        /// Deletes all eligible contents from a destination view.
        /// Mirrors pyRevit clear_view_contents.
        /// </summary>
        public static void ClearViewContents(
            Document destDoc, View destView,
            SheetCopyOptions options)
        {
            var ids = GetViewContents(destDoc, destView, options);

            foreach (var id in ids)
            {
                try { destDoc.Delete(id); }
                catch { /* element may be protected */ }
            }
        }

        /// <summary>
        /// Copies all view contents from source view to destination
        /// view using the VIEW-TO-VIEW overload of CopyElements.
        /// This is the key method that preserves dimensions, spot
        /// elevations, text notes, detail lines, etc.
        /// Mirrors pyRevit copy_view_contents.
        /// </summary>
        public static void CopyViewContentsViewToView(
            Document sourceDoc, View sourceView,
            Document destDoc, View destView,
            SheetCopyOptions options)
        {
            var ids = GetViewContents(sourceDoc, sourceView, options);
            if (ids.Count == 0) return;

            var opts = new CopyPasteOptions();
            opts.SetDuplicateTypeNamesHandler(
                new UseDestinationTypesHandler());

            // VIEW-TO-VIEW overload — this is what preserves
            // all view-specific content (dimensions, spots, etc.)
            ElementTransformUtils.CopyElements(
                sourceView,
                ids,
                destView,
                null,   // no additional transform
                opts);
        }

        /// <summary>
        /// Copies scale and view description from source to dest.
        /// Mirrors pyRevit copy_view_props.
        /// </summary>
        public static void CopyViewProperties(
            View sourceView, View destView)
        {
            // Scale
            try { destView.Scale = sourceView.Scale; } catch { }

            // View Description
            try
            {
                var srcParam = sourceView.get_Parameter(
                    BuiltInParameter.VIEW_DESCRIPTION);
                var dstParam = destView.get_Parameter(
                    BuiltInParameter.VIEW_DESCRIPTION);
                if (srcParam != null && dstParam != null
                    && !dstParam.IsReadOnly)
                {
                    dstParam.Set(srcParam.AsString() ?? "");
                }
            }
            catch { }

            // Detail Level (Coarse / Medium / Fine)
            try
            {
                destView.DetailLevel = sourceView.DetailLevel;
            }
            catch { }

            // Visual Style (Wireframe / Hidden / Shaded / etc.)
            try
            {
                destView.DisplayStyle = sourceView.DisplayStyle;
            }
            catch { }

            // Full CropBox (includes rotation for rotated plans)
            try
            {
                if (sourceView.CropBoxActive)
                {
                    destView.CropBox = sourceView.CropBox;
                    destView.CropBoxActive = true;
                    destView.CropBoxVisible = sourceView.CropBoxVisible;
                }
            }
            catch { }
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
                $"[DBG-SH1] CopySheets called with {sheetIds.Count} sheet(s).");
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
                result.Warnings.Add(
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
            ViewSheet destSheet = FindMatchingView(
                target, srcSheet) as ViewSheet;
            result.Warnings.Add(
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
                        ClearViewContents(target, destSheet, opts);
                        CopyViewContentsViewToView(
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
                result.Warnings.Add(
                    $"[DBG-SH4] '{srcSheet.SheetNumber}': created destSheet id={destSheet.Id.IntegerValue}");
                // Copy sheet-level content (title block, schedules)
                // in a SEPARATE transaction after sheet creation
                using (var t = new Transaction(target,
                    "HMV – Copy Sheet Contents"))
                {
                    t.Start();
                    try
                    {
                        CopyViewContentsViewToView(
                            source, srcSheet, target, destSheet, opts);
                        t.Commit();
                        result.Warnings.Add(
                            $"[DBG-SH5] '{srcSheet.SheetNumber}': sheet content copied OK.");
                    }
                    catch (Exception ex)
                    {
                        t.RollBack();
                        result.Warnings.Add(
                            $"[DBG-SH5] '{srcSheet.SheetNumber}': sheet content FAILED: {ex.Message}");
                    }
                }
            }

            if (destSheet == null) return null;

            // ── Viewports ─────────────────────────────────────
            // Always attempt viewport copy for non-placeholder sheets
            result.Warnings.Add(
                $"[DBG-SH6] '{srcSheet.SheetNumber}': "
                + $"isPlaceholder={destSheet.IsPlaceholder} "
                + $"CopyViewports={opts.CopyViewports} "
                + $"srcViewportCount={srcSheet.GetAllViewports().Count}");
            if (!destSheet.IsPlaceholder && opts.CopyViewports)
            {
                try
                {
                    CopySheetViewports(
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
                CopyViewAnnotations(source, target, viewMap, result);
                result.Warnings.Add(
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
                CopySheetRevisions(
                    source, srcSheet, target, destSheet, result);
            }

            return destSheet;
        }




        // ══════════════════════════════════════════════════════
        //  7a. COPY SHEET VIEWPORTS
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Iterates source sheet viewports, copies/updates the
        /// referenced view, then places it on the destination sheet
        /// at the original position with matching viewport type,
        /// detail number, label offset, and bounding-box correction.
        /// Mirrors pyRevit copy_sheet_viewports.
        /// </summary>
        private static void CopySheetViewports(
            Document source, ViewSheet srcSheet,
            Document target, ViewSheet destSheet,
            Transform coordTransform,
            Dictionary<ElementId, ElementId> viewMap,
            TransferResult result,
            MigrationSettings settings)
        {
            var existingViewIds = new HashSet<ElementId>(
                destSheet.GetAllViewports()
                    .Select(vpId => target.GetElement(vpId))
                    .OfType<Viewport>()
                    .Select(vp => vp.ViewId));

            foreach (var srcVportId in srcSheet.GetAllViewports())
            {
                Viewport srcVport = source.GetElement(srcVportId)
                    as Viewport;
                if (srcVport == null) continue;

                View srcView = source.GetElement(srcVport.ViewId)
                    as View;
                if (srcView == null) continue;

                result.Warnings.Add(
                    $"[DBG-SH7] Sheet '{srcSheet.SheetNumber}' viewport: "
                    + $"view='{srcView.Name}' type={srcView.ViewType} "
                    + $"srcViewId={srcVport.ViewId.IntegerValue} "
                    + $"alreadyInMap={viewMap.ContainsKey(srcVport.ViewId)}");

                // ── Copy or update the referenced view ────────
                ElementId destViewId;
                if (viewMap.ContainsKey(srcVport.ViewId))
                {
                    destViewId = viewMap[srcVport.ViewId];
                }
                else
                {
                    destViewId = CopyOrUpdateSingleView(
                        source, target, srcView,
                        coordTransform, result, settings);

                    if (destViewId != ElementId.InvalidElementId)
                        viewMap[srcVport.ViewId] = destViewId;
                }
                result.Warnings.Add(
                    $"[DBG-SH8] '{srcView.Name}': "
                    + $"destViewId={(destViewId != ElementId.InvalidElementId ? destViewId.IntegerValue.ToString() : "INVALID")}");
                if (destViewId == ElementId.InvalidElementId)
                {
                    result.Warnings.Add(
                        $"Could not copy view '{srcView.Name}' "
                        + "for viewport placement.");
                    continue;
                }

                View destView = target.GetElement(destViewId) as View;
                if (destView == null) continue;

                // Skip if already on this sheet
                if (existingViewIds.Contains(destViewId))
                {
                    result.Warnings.Add(
                        $"View '{srcView.Name}' already on sheet "
                        + $"'{destSheet.SheetNumber}'.");
                    continue;
                }

                // Check if placed on a different sheet
                string placedSheet = GetViewSheetNumber(
                    target, destView);
                if (placedSheet != null
                    && placedSheet != destSheet.SheetNumber)
                {
                    result.Warnings.Add(
                        $"View '{srcView.Name}' already placed on "
                        + $"sheet '{placedSheet}', skipping.");
                    continue;
                }

                // ── Capture source viewport data ──────────────
                var srcData = CaptureViewportData(
                    source, srcVport, srcSheet);

                // ── Place viewport ────────────────────────────
                Viewport newVport = null;
                using (var t = new Transaction(target,
                    "HMV – Place Viewport"))
                {
                    t.Start();
                    try
                    {
                        newVport = Viewport.Create(
                            target,
                            destSheet.Id,
                            destViewId,
                            srcData.BoxCenter);

                        // Apply rotation BEFORE position correction
                        if (srcData.Rotation != ViewportRotation.None)
                        {
                            try
                            {
                                newVport.Rotation = srcData.Rotation;
                            }
                            catch (Exception rex)
                            {
                                result.Warnings.Add(
                                    $"Could not rotate viewport for "
                                    + $"'{srcView.Name}': {rex.Message}");
                            }
                        }

                        // Preserve detail number
                        if (settings.SheetOptions.PreserveDetailNumbers)
                        {
                            ApplyDetailNumber(
                                srcVport, newVport, result);
                        }

                        t.Commit();
                        result.ViewportsCopied++;
                        result.Warnings.Add(
                            $"[DBG-SH9] '{srcView.Name}': viewport placed OK "
                            + $"at ({srcData.BoxCenter.X:F4},{srcData.BoxCenter.Y:F4}) "
                            + $"rotation={srcData.Rotation} vpId={newVport.Id.IntegerValue}");
                    }
                    catch (Exception ex)
                    {
                        t.RollBack();
                        result.Warnings.Add(
                            $"Could not place viewport for "
                            + $"'{srcView.Name}': {ex.Message}");
                        continue;
                    }
                }

                if (newVport == null) continue;

                // ── Match viewport type ───────────────────────
                ApplyViewportType(
                    source, srcVportId, target, newVport.Id, result);

                // ── Label offset and line length (2022+) ──────
                ApplyLabelProperties(
                    target, newVport.Id, srcData, result);

                // ── BBox position correction (AFTER rotation) ─
                CorrectViewportByBBox(
                    target, newVport.Id, srcData, destSheet, result);
            }
        }



        // ═══════════════════════════════════════════════════════════════
        //  REPLACE: CopyOrUpdateSingleView
        //  Search: "private static ElementId CopyOrUpdateSingleView"
        //  DELETE everything from that line until the closing brace.
        //
        //  ALSO DELETE these methods (no longer used):
        //    - CopyViewDocToDoc
        //    - TryForceCreateFallback
        //    - CopyDraftingViewDocToDoc
        //    - CopyLegendDocToDoc
        //    - CreateDraftingView (old)
        //    - CreateLegendView (old)
        //
        //  KEEP these methods (still used as fallbacks):
        //    - ForceCreateView
        //    - ForceCreatePlan
        //    - ForceCreateSection
        //    - FindViewFamilyTypeId
        //
        //  ADD: SwallowErrorsHandler class (outside TransferManager,
        //       inside namespace HMVTools, after ViewportData class)
        // ═══════════════════════════════════════════════════════════════


        // ───────────────────────────────────────────────────────
        //  PASTE THIS: CopyOrUpdateSingleView
        //  (exact pyRevit copy_view pattern)
        // ───────────────────────────────────────────────────────

        private static ElementId CopyOrUpdateSingleView(
            Document source, Document target,
            View srcView,
            Transform coordTransform,
            TransferResult result,
            MigrationSettings settings)
        {
            ViewTransferMode mode = settings.TransferMode;
            var opts = settings.SheetOptions;

            // ── 1. CHECK FOR EXISTING VIEW (pyRevit: find_matching_view) ──
            View matchingView = FindMatchingView(target, srcView);

            if (matchingView != null)
            {
                if (mode == ViewTransferMode.Update)
                {
                    // pyRevit: clear_view_contents + copy_view_contents
                    using (var t = new Transaction(target,
                        "HMV – Update View Contents"))
                    {
                        SetSwallowErrors(t);
                        t.Start();
                        ClearViewContents(target, matchingView, opts);
                        CopyViewContentsViewToView(
                            source, srcView, target, matchingView, opts);
                        CopyViewProperties(srcView, matchingView);
                        t.Commit();
                    }
                    result.ViewsUpdated++;
                }
                return matchingView.Id;
            }

            // ── 2. CREATE NEW VIEW (pyRevit: create based on type) ────
            View newView = null;

            if (srcView.ViewType == ViewType.DrawingSheet)
            {
                // Sheets handled separately
                return ElementId.InvalidElementId;
            }
            else if (srcView.ViewType == ViewType.DraftingView)
            {
                result.Warnings.Add(
                    $"[DEBUG] Drafting '{srcView.Name}': starting doc-to-doc copy.");

                // Snapshot existing view IDs to find the new one
                var existingIds = new HashSet<int>(
                    new FilteredElementCollector(target)
                        .OfClass(typeof(View))
                        .Select(v => v.Id.IntegerValue));

                result.Warnings.Add(
                    $"[DEBUG] Drafting '{srcView.Name}': "
                    + $"target has {existingIds.Count} existing views before copy.");

                // Check source view contents BEFORE copying
                var srcContents = new FilteredElementCollector(source, srcView.Id)
                    .WhereElementIsNotElementType()
                    .ToElementIds();
                result.Warnings.Add(
                    $"[DEBUG] Drafting '{srcView.Name}': "
                    + $"source view has {srcContents.Count} elements.");

                using (var t = new Transaction(target,
                    "HMV – Copy Drafting View"))
                {
                    t.Start();
                    try
                    {
                        var cpOpts = new CopyPasteOptions();
                        cpOpts.SetDuplicateTypeNamesHandler(
                            new UseDestinationTypesHandler());

                        var copied = ElementTransformUtils.CopyElements(
                            source,
                            new List<ElementId> { srcView.Id },
                            target,
                            Transform.Identity,
                            cpOpts);

                        result.Warnings.Add(
                            $"[DEBUG] Drafting '{srcView.Name}': "
                            + $"CopyElements returned {copied?.Count ?? 0} ids.");

                        if (copied != null && copied.Count > 0)
                        {
                            foreach (var cid in copied)
                            {
                                Element el = target.GetElement(cid);
                                bool isNew = !existingIds.Contains(cid.IntegerValue);
                                result.Warnings.Add(
                                    $"[DEBUG]   id={cid.IntegerValue} "
                                    + $"type={el?.GetType().Name ?? "null"} "
                                    + $"isNew={isNew}");
                            }

                            ElementId newId = copied.FirstOrDefault(id =>
                                !existingIds.Contains(id.IntegerValue)
                                && target.GetElement(id) is View);

                            if (newId != null
                                && newId != ElementId.InvalidElementId)
                            {
                                newView = target.GetElement(newId) as View;

                                // Check contents of new view BEFORE commit
                                var newContents = new FilteredElementCollector(
                                        target, newId)
                                    .WhereElementIsNotElementType()
                                    .ToElementIds();
                                result.Warnings.Add(
                                    $"[DEBUG] Drafting '{srcView.Name}': "
                                    + $"new view has {newContents.Count} elements before commit.");

                                t.Commit();
                                result.Warnings.Add(
                                    $"[DEBUG] Drafting '{srcView.Name}': committed OK.");
                            }
                            else
                            {
                                t.RollBack();
                                result.Warnings.Add(
                                    $"[DEBUG] Drafting '{srcView.Name}': "
                                    + "no new View found in copied ids — rolled back.");
                            }
                        }
                        else
                        {
                            t.RollBack();
                            result.Warnings.Add(
                                $"[DEBUG] Drafting '{srcView.Name}': "
                                + "CopyElements returned null or empty — rolled back.");
                        }
                    }
                    catch (Exception ex)
                    {
                        t.RollBack();
                        result.Warnings.Add(
                            $"[DEBUG] Drafting '{srcView.Name}' EXCEPTION: "
                            + $"{ex.GetType().Name}: {ex.Message}");
                    }
                }

                // After transaction — check final state
                if (newView != null)
                {
                    var finalContents = new FilteredElementCollector(
                            target, newView.Id)
                        .WhereElementIsNotElementType()
                        .ToElementIds();
                    result.Warnings.Add(
                        $"[DEBUG] Drafting '{srcView.Name}': "
                        + $"final view '{newView.Name}' has {finalContents.Count} elements after commit.");
                }
                else
                {
                    result.Warnings.Add(
                        $"[DEBUG] Drafting '{srcView.Name}': newView is null after transaction.");
                }
            }
            else if (srcView.ViewType == ViewType.Legend)
            {
                // pyRevit: duplicate first legend
                using (var t = new Transaction(target,
                    "HMV – Create Legend View"))
                {
                    t.Start();
                    try
                    {
                        View firstLegend =
                            new FilteredElementCollector(target)
                                .OfClass(typeof(View))
                                .Cast<View>()
                                .FirstOrDefault(v =>
                                    v.ViewType == ViewType.Legend
                                    && !v.IsTemplate);

                        if (firstLegend != null)
                        {
                            ElementId newId = firstLegend.Duplicate(
                                ViewDuplicateOption.Duplicate);
                            newView = target.GetElement(newId) as View;
                            if (newView != null)
                            {
                                newView.Name = GetUniqueViewName(
                                    target, srcView.Name);
                                CopyViewProperties(srcView, newView);
                            }
                        }
                        else
                        {
                            result.Warnings.Add(
                                "No legend in target to duplicate.");
                        }
                        t.Commit();
                    }
                    catch (Exception ex)
                    {
                        t.RollBack();
                        result.Warnings.Add(
                            $"Error creating legend "
                            + $"'{srcView.Name}': {ex.Message}");
                    }
                }
            }
            else if (srcView.ViewType == ViewType.FloorPlan
                  || srcView.ViewType == ViewType.CeilingPlan)
            {
                // Use doc-to-doc CopyElements (same as CopyViews PASS 1)
                using (var t = new Transaction(target,
                    "HMV – Copy Plan Doc-to-Doc"))
                {
                    t.Start();
                    try
                    {
                        var existingIds = new HashSet<int>(
                            new FilteredElementCollector(target)
                                .OfClass(typeof(View))
                                .Select(v => v.Id.IntegerValue));

                        var cpOpts = new CopyPasteOptions();
                        cpOpts.SetDuplicateTypeNamesHandler(
                            new UseDestinationTypesHandler());

                        var copied = ElementTransformUtils.CopyElements(
                            source,
                            new List<ElementId> { srcView.Id },
                            target,
                            coordTransform,
                            cpOpts);

                        if (copied != null && copied.Count > 0)
                        {
                            ElementId newId = copied.FirstOrDefault(id =>
                                !existingIds.Contains(id.IntegerValue)
                                && target.GetElement(id) is View);

                            if (newId != null
                                && newId != ElementId.InvalidElementId)
                            {
                                newView = target.GetElement(newId) as View;
                                t.Commit();
                                result.Warnings.Add(
                                    $"[DBG-PLAN] '{srcView.Name}': "
                                    + $"doc-to-doc copy OK, id={newId.IntegerValue}");
                            }
                            else
                            {
                                t.RollBack();
                                result.Warnings.Add(
                                    $"[DBG-PLAN] '{srcView.Name}': "
                                    + "no new View found in copied ids.");

                                // Fallback to ForceCreatePlan
                                ElementId fallbackId = ForceCreatePlan(
                                    source, target,
                                    srcView as ViewPlan, result);
                                if (fallbackId != ElementId.InvalidElementId)
                                    newView = target.GetElement(fallbackId) as View;
                            }
                        }
                        else
                        {
                            t.RollBack();
                            result.Warnings.Add(
                                $"[DBG-PLAN] '{srcView.Name}': "
                                + "CopyElements returned empty, trying ForceCreatePlan.");

                            ElementId fallbackId = ForceCreatePlan(
                                source, target,
                                srcView as ViewPlan, result);
                            if (fallbackId != ElementId.InvalidElementId)
                                newView = target.GetElement(fallbackId) as View;
                        }
                    }
                    catch (Exception ex)
                    {
                        if (t.HasStarted() && !t.HasEnded())
                            t.RollBack();

                        result.Warnings.Add(
                            $"[DBG-PLAN] '{srcView.Name}': "
                            + $"doc-to-doc failed: {ex.Message}, trying ForceCreatePlan.");

                        ElementId fallbackId = ForceCreatePlan(
                            source, target,
                            srcView as ViewPlan, result);
                        if (fallbackId != ElementId.InvalidElementId)
                            newView = target.GetElement(fallbackId) as View;
                    }
                }
            }
            else if (srcView.ViewType == ViewType.Section
                  || srcView.ViewType == ViewType.Detail
                  || srcView.ViewType == ViewType.Elevation)
            {
                // ForceCreateSection manages its own Transactions internally
                ElementId id = ForceCreateSection(
                    source, target,
                    srcView as ViewSection, result);
                if (id != ElementId.InvalidElementId)
                    newView = target.GetElement(id) as View;
            }
            if (newView == null)
                return ElementId.InvalidElementId;

            result.ViewsCreated++;

            // ── 3. COPY VIEW CONTENTS (pyRevit: copy_view_contents) ──
            // ForceCreatePlan uses Duplicate(WithDetailing) — content
            // is already included. ForceCreateSection copies content
            // in its own STEP 6. Only copy content for view types
            // that DON'T handle their own content above.
            bool contentAlreadyCopied =
                  srcView.ViewType == ViewType.FloorPlan
               || srcView.ViewType == ViewType.CeilingPlan
               || srcView.ViewType == ViewType.Section
               || srcView.ViewType == ViewType.Detail
               || srcView.ViewType == ViewType.Elevation;

            if (!contentAlreadyCopied)
            {
                using (var t = new Transaction(target,
                    "HMV – Copy View Contents"))
                {
                    SetSwallowErrors(t);
                    t.Start();

                    try
                    {
                        var contentIds = GetViewContents(
                            source, srcView, opts);

                        result.Warnings.Add(
                            $"[DBG-CONTENT] '{srcView.Name}': "
                            + $"GetViewContents returned {contentIds.Count} elements.");

                        if (contentIds.Count > 0)
                        {
                            // Log element types being copied
                            var typeCounts = contentIds
                                .Select(id => source.GetElement(id))
                                .Where(e => e != null)
                                .GroupBy(e => e.GetType().Name)
                                .Select(g => $"{g.Key}={g.Count()}")
                                .ToArray();
                            result.Warnings.Add(
                                $"[DBG-CONTENT] '{srcView.Name}': types: "
                                + string.Join(", ", typeCounts));

                            var cpOpts = new CopyPasteOptions();
                            cpOpts.SetDuplicateTypeNamesHandler(
                                new UseDestinationTypesHandler());
                            // Only keep view-owned elements (annotations,
                            // dims, text, detail items). Skip model elements
                            // that are merely visible in the view.
                            contentIds = contentIds.Where(id =>
                            {
                                Element e = source.GetElement(id);
                                if (e == null) return false;
                                if (e is SketchPlane) return false;
                                if (e.OwnerViewId == ElementId.InvalidElementId)
                                    return false;
                                return true;
                            }).ToList();

                            result.Warnings.Add(
                                $"[DBG-CONTENT] '{srcView.Name}': "
                                + $"after filter (view-owned only): {contentIds.Count} elements.");

                            var copiedIds = ElementTransformUtils.CopyElements(
                                srcView,
                                contentIds,
                                newView,
                                null,
                                cpOpts);

                            result.Warnings.Add(
                                $"[DBG-CONTENT] '{srcView.Name}': "
                                + $"CopyElements returned {copiedIds?.Count ?? 0} elements.");
                        }
                    }
                    catch (Exception ex)
                    {
                        result.Warnings.Add(
                            $"[DBG-CONTENT] '{srcView.Name}': FAILED: "
                            + $"{ex.GetType().Name}: {ex.Message}");
                    }

                    if (t.HasStarted() && !t.HasEnded())
                    {
                        try { t.Commit(); }
                        catch { try { t.RollBack(); } catch { } }
                    }
                }
            }
            else
            {
                result.Warnings.Add(
                    $"[DBG-COSV] '{srcView.Name}': skipped content copy "
                    + $"(handled by ForceCreate, type={srcView.ViewType})");
            }

            return newView.Id;

        }

        /// <summary>
        /// Sets failure handling to swallow all warnings and
        /// continue on non-fatal errors. Equivalent to pyRevit's
        /// swallow_errors=True.
        /// </summary>
        private static void SetSwallowErrors(Transaction t)
        {
            var opts = t.GetFailureHandlingOptions();
            opts.SetFailuresPreprocessor(
                new SwallowErrorsHandler());
            opts.SetClearAfterRollback(true);
            t.SetFailureHandlingOptions(opts);
        }


        /// <summary>
        /// Creates a view in the target using doc-to-doc CopyElements.
        /// This is equivalent to copy/paste-aligned — preserves exact
        /// geometry for sections, drafting views, legends, etc.
        /// Returns the new view Id, or InvalidElementId if it fails.
        /// </summary>
        private static ElementId CopyViewDocToDoc(
            Document source, Document target,
            View srcView, Transform coordTransform,
            TransferResult result)
        {
            var cpOpts = new CopyPasteOptions();
            cpOpts.SetDuplicateTypeNamesHandler(
                new UseDestinationTypesHandler());

            // Choose transform: identity for flat views,
            // coordTransform for spatial views
            Transform xform = Transform.Identity;
            switch (srcView.ViewType)
            {
                case ViewType.FloorPlan:
                case ViewType.CeilingPlan:
                case ViewType.Section:
                case ViewType.Detail:
                case ViewType.Elevation:
                    // For spatial views with shared coords,
                    // check if transform is near-identity
                    if (coordTransform != null)
                        xform = coordTransform;
                    break;
            }

            // Snapshot existing view IDs
            var existingIds = new HashSet<int>(
                new FilteredElementCollector(target)
                    .OfClass(typeof(View))
                    .Select(v => v.Id.IntegerValue));

            using (var t = new Transaction(target,
                "HMV – Copy View Doc-to-Doc"))
            {
                t.Start();
                try
                {
                    var copied = ElementTransformUtils.CopyElements(
                        source,
                        new List<ElementId> { srcView.Id },
                        target,
                        xform,
                        cpOpts);

                    if (copied != null && copied.Count > 0)
                    {
                        // Find the NEW view (not in the snapshot)
                        ElementId newId = null;
                        foreach (var id in copied)
                        {
                            if (!existingIds.Contains(id.IntegerValue))
                            {
                                Element el = target.GetElement(id);
                                if (el is View)
                                {
                                    newId = id;
                                    break;
                                }
                            }
                        }

                        if (newId != null)
                        {
                            t.Commit();
                            return newId;
                        }

                        // CopyElements returned only existing IDs —
                        // Revit reused an existing view instead of
                        // creating a new one. Roll back.
                        t.RollBack();
                        result.Warnings.Add(
                            $"Doc-to-doc copy of '{srcView.Name}' "
                            + "returned existing view, trying "
                            + "fallback.");
                        return ElementId.InvalidElementId;
                    }

                    t.RollBack();
                    return ElementId.InvalidElementId;
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    result.Warnings.Add(
                        $"Doc-to-doc copy of '{srcView.Name}' "
                        + $"failed: {ex.Message}");
                    return ElementId.InvalidElementId;
                }
            }
        }
        /// <summary>
        /// Fallback when doc-to-doc fails: uses ForceCreatePlan or
        /// ForceCreateSection to reconstruct the view from geometry.
        /// Only for spatial views — drafting/legend have no fallback.
        /// </summary>
        private static ElementId TryForceCreateFallback(
            Document source, Document target,
            View srcView, Transform coordTransform,
            TransferResult result)
        {
            ElementId newId = ElementId.InvalidElementId;

            using (var t = new Transaction(target,
                "HMV – ForceCreate Fallback"))
            {
                t.Start();
                try
                {
                    switch (srcView.ViewType)
                    {
                        case ViewType.FloorPlan:
                        case ViewType.CeilingPlan:
                            newId = ForceCreatePlan(
                                source, target,
                                srcView as ViewPlan, result);
                            break;

                        case ViewType.Section:
                        case ViewType.Detail:
                            newId = ForceCreateSection(
                                source, target,
                                srcView as ViewSection, result);
                            break;

                        case ViewType.Elevation:
                            result.Warnings.Add(
                                $"Elevation '{srcView.Name}' "
                                + "cannot be created. Use Update "
                                + "mode or place manually.");
                            break;

                        // Drafting/Legend: no fallback — doc-to-doc
                        // is the only path.
                        default:
                            break;
                    }

                    if (newId != null
                        && newId != ElementId.InvalidElementId)
                    {
                        t.Commit();
                        return newId;
                    }

                    t.RollBack();
                    return ElementId.InvalidElementId;
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    result.Warnings.Add(
                        $"Fallback create failed for "
                        + $"'{srcView.Name}': {ex.Message}");
                    return ElementId.InvalidElementId;
                }
            }
        }


        /// <summary>
        /// Copies a DraftingView using doc-to-doc CopyElements.
        /// This brings the view AND all its contents in one shot.
        /// </summary>
        private static ElementId CopyDraftingViewDocToDoc(
            Document source, Document target,
            View srcView, TransferResult result)
        {
            var cpOpts = new CopyPasteOptions();
            cpOpts.SetDuplicateTypeNamesHandler(
                new UseDestinationTypesHandler());

            // Snapshot existing view IDs to find the new one
            var existingIds = new HashSet<int>(
                new FilteredElementCollector(target)
                    .OfClass(typeof(View))
                    .Select(v => v.Id.IntegerValue));

            using (var t = new Transaction(target,
                "HMV – Copy Drafting View"))
            {
                t.Start();
                try
                {
                    var copied = ElementTransformUtils.CopyElements(
                        source,
                        new List<ElementId> { srcView.Id },
                        target,
                        Transform.Identity,
                        cpOpts);

                    if (copied != null && copied.Count > 0)
                    {
                        // Find the new view (not in the snapshot)
                        ElementId newId = copied
                            .FirstOrDefault(id =>
                                !existingIds.Contains(id.IntegerValue)
                                && target.GetElement(id) is View);

                        if (newId == null)
                            newId = copied.First();

                        t.Commit();
                        result.ViewsCreated++;
                        return newId;
                    }

                    t.RollBack();
                    result.Warnings.Add(
                        $"Drafting view '{srcView.Name}' copy "
                        + "returned no elements.");
                    return ElementId.InvalidElementId;
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    result.Warnings.Add(
                        $"Drafting view '{srcView.Name}' copy "
                        + $"failed: {ex.Message}");
                    return ElementId.InvalidElementId;
                }
            }
        }

        /// <summary>
        /// Copies a Legend view using doc-to-doc CopyElements.
        /// </summary>
        private static ElementId CopyLegendDocToDoc(
            Document source, Document target,
            View srcView, TransferResult result)
        {
            var cpOpts = new CopyPasteOptions();
            cpOpts.SetDuplicateTypeNamesHandler(
                new UseDestinationTypesHandler());

            var existingIds = new HashSet<int>(
                new FilteredElementCollector(target)
                    .OfClass(typeof(View))
                    .Select(v => v.Id.IntegerValue));

            using (var t = new Transaction(target,
                "HMV – Copy Legend View"))
            {
                t.Start();
                try
                {
                    var copied = ElementTransformUtils.CopyElements(
                        source,
                        new List<ElementId> { srcView.Id },
                        target,
                        Transform.Identity,
                        cpOpts);

                    if (copied != null && copied.Count > 0)
                    {
                        ElementId newId = copied
                            .FirstOrDefault(id =>
                                !existingIds.Contains(id.IntegerValue)
                                && target.GetElement(id) is View);

                        if (newId == null)
                            newId = copied.First();

                        t.Commit();
                        result.ViewsCreated++;
                        return newId;
                    }

                    t.RollBack();
                    result.Warnings.Add(
                        $"Legend '{srcView.Name}' copy "
                        + "returned no elements.");
                    return ElementId.InvalidElementId;
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    result.Warnings.Add(
                        $"Legend '{srcView.Name}' copy "
                        + $"failed: {ex.Message}");
                    return ElementId.InvalidElementId;
                }
            }
        }




        /// <summary>
        /// Creates a DraftingView in the target document.
        /// Mirrors pyRevit's drafting view creation.
        /// </summary>
        private static ElementId CreateDraftingView(
            Document source, Document target,
            View srcView, TransferResult result)
        {
            ElementId typeId = target.GetDefaultElementTypeId(
                ElementTypeGroup.ViewTypeDrafting);

            if (typeId == null || typeId == ElementId.InvalidElementId)
            {
                result.Warnings.Add(
                    "No default drafting view type in target.");
                return ElementId.InvalidElementId;
            }

            ViewDrafting newView = ViewDrafting.Create(target, typeId);
            if (newView == null) return ElementId.InvalidElementId;

            newView.Name = GetUniqueViewName(target, srcView.Name);
            CopyViewProperties(srcView, newView);

            return newView.Id;
        }

        /// <summary>
        /// Creates a Legend view in the target by duplicating the
        /// first existing legend. Mirrors pyRevit's legend creation.
        /// </summary>
        private static ElementId CreateLegendView(
            Document source, Document target,
            View srcView, TransferResult result)
        {
            // Find first legend in target to duplicate
            View firstLegend = new FilteredElementCollector(target)
                .OfClass(typeof(View))
                .Cast<View>()
                .FirstOrDefault(v => v.ViewType == ViewType.Legend
                    && !v.IsTemplate);

            if (firstLegend == null)
            {
                result.Warnings.Add(
                    "Target document has no Legend view to "
                    + "duplicate. Skipping legend.");
                return ElementId.InvalidElementId;
            }

            ElementId newId = firstLegend.Duplicate(
                ViewDuplicateOption.Duplicate);

            View newView = target.GetElement(newId) as View;
            if (newView == null) return ElementId.InvalidElementId;

            newView.Name = GetUniqueViewName(target, srcView.Name);
            CopyViewProperties(srcView, newView);

            return newView.Id;
        }


        // ═══════════════════════════════════════════════════════
        //  7c. VIEWPORT TYPE MATCHING
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Copies the viewport type from source to target if needed,
        /// then applies it to the new viewport.
        /// Mirrors pyRevit apply_viewport_type + copy_viewport_types.
        /// </summary>
        private static void ApplyViewportType(
            Document source, ElementId srcVportId,
            Document target, ElementId destVportId,
            TransferResult result)
        {
            using (var t = new Transaction(target,
                "HMV – Apply Viewport Type"))
            {
                t.Start();
                try
                {
                    Viewport srcVport = source.GetElement(srcVportId)
                        as Viewport;
                    Element srcType = source.GetElement(
                        srcVport.GetTypeId());
                    string srcTypeName = srcType.Name;

                    Viewport destVport = target.GetElement(destVportId)
                        as Viewport;

                    // Check if type name exists in target
                    var destTypeNames = new HashSet<string>(
                        destVport.GetValidTypes()
                            .Select(id => target.GetElement(id).Name));

                    if (!destTypeNames.Contains(srcTypeName))
                    {
                        // Copy the viewport type to target
                        var cpOpts = new CopyPasteOptions();
                        cpOpts.SetDuplicateTypeNamesHandler(
                            new UseDestinationTypesHandler());

                        ElementTransformUtils.CopyElements(
                            source,
                            new List<ElementId> { srcType.Id },
                            target,
                            null,
                            cpOpts);
                    }

                    // Find and apply matching type
                    foreach (var vtId in destVport.GetValidTypes())
                    {
                        Element vt = target.GetElement(vtId);
                        if (vt.Name == srcTypeName)
                        {
                            destVport.ChangeTypeId(vtId);
                            break;
                        }
                    }

                    t.Commit();
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    result.Warnings.Add(
                        $"Could not apply viewport type: "
                        + ex.Message);
                }
            }
        }

        /// <summary>
        /// Applies detail number from source to destination viewport.
        /// Mirrors pyRevit apply_detail_number.
        /// </summary>
        private static void ApplyDetailNumber(
            Viewport srcVport, Viewport destVport,
            TransferResult result)
        {
            try
            {
                var srcParam = srcVport.get_Parameter(
                    BuiltInParameter.VIEWPORT_DETAIL_NUMBER);
                if (srcParam == null) return;

                string detailNum = srcParam.AsString();
                if (string.IsNullOrEmpty(detailNum)) return;

                var destParam = destVport.get_Parameter(
                    BuiltInParameter.VIEWPORT_DETAIL_NUMBER);
                if (destParam != null && !destParam.IsReadOnly)
                {
                    destParam.Set(detailNum);
                }
            }
            catch (Exception ex)
            {
                result.Warnings.Add(
                    $"Could not preserve detail number: "
                    + ex.Message);
            }
        }


        // ═══════════════════════════════════════════════════════
        //  7d. VIEWPORT POSITIONING (Label, BBox correction)
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Captures source viewport placement data: center, label
        /// offset, label line length, and bounding box on the sheet.
        /// Mirrors pyRevit get_source_vport_data.
        /// </summary>
        public static ViewportData CaptureViewportData(
             Document doc, Viewport vport, ViewSheet sheet)
        {
            var data = new ViewportData
            {
                BoxCenter = vport.GetBoxCenter()
            };

            // Rotation
            try { data.Rotation = vport.Rotation; }
            catch { data.Rotation = ViewportRotation.None; }

            // Label offset (Revit 2022+)
            try { data.LabelOffset = vport.LabelOffset; }
            catch { }

            // Label line length (Revit 2022+)
            try { data.LabelLineLength = vport.LabelLineLength; }
            catch { }

            // Bounding box on sheet (includes view title)
            try
            {
                BoundingBoxXYZ bb = vport.get_BoundingBox(sheet);
                if (bb != null)
                {
                    data.BBoxMin = bb.Min;
                    data.BBoxMax = bb.Max;
                }
            }
            catch { }

            return data;
        }


        /// <summary>
        /// Sets label offset and line length on destination viewport.
        /// Mirrors pyRevit apply_vport_label_props.
        /// </summary>
        private static void ApplyLabelProperties(
            Document destDoc, ElementId destVportId,
            ViewportData srcData, TransferResult result)
        {
            if (srcData.LabelOffset == null
                && srcData.LabelLineLength == null)
                return;

            using (var t = new Transaction(destDoc,
                "HMV – Set View Title Properties"))
            {
                t.Start();
                try
                {
                    Viewport vp = destDoc.GetElement(destVportId)
                        as Viewport;

                    if (srcData.LabelOffset != null)
                    {
                        try { vp.LabelOffset = srcData.LabelOffset; }
                        catch { /* 2022+ */ }
                    }
                    if (srcData.LabelLineLength != null)
                    {
                        try
                        {
                            vp.LabelLineLength =
                                srcData.LabelLineLength.Value;
                        }
                        catch { /* 2022+ */ }
                    }

                    t.Commit();
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    result.Warnings.Add(
                        $"Label property set failed: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Verifies viewport position using BoundingBox(sheet) and
        /// corrects any residual drift with MoveElement.
        /// Mirrors pyRevit correct_vport_by_bbox.
        /// </summary>
        private static void CorrectViewportByBBox(
            Document destDoc, ElementId destVportId,
            ViewportData srcData, ViewSheet destSheet,
            TransferResult result)
        {
            if (srcData.BBoxMin == null) return;

            try
            {
                Viewport vp = destDoc.GetElement(destVportId)
                    as Viewport;
                BoundingBoxXYZ dstBb = vp.get_BoundingBox(destSheet);
                if (dstBb == null) return;

                double dx = srcData.BBoxMin.X - dstBb.Min.X;
                double dy = srcData.BBoxMin.Y - dstBb.Min.Y;

                if (Math.Abs(dx) <= 1e-9 && Math.Abs(dy) <= 1e-9)
                    return;

                using (var t = new Transaction(destDoc,
                    "HMV – Align Viewport Position"))
                {
                    t.Start();
                    result.Warnings.Add(
                        $"[DBG-SH10] BBox correction: "
                        + $"srcBBMin=({srcData.BBoxMin.X:F4},{srcData.BBoxMin.Y:F4}) "
                        + $"dstBBMin=({dstBb.Min.X:F4},{dstBb.Min.Y:F4}) "
                        + $"dx={dx:F6} dy={dy:F6}");
                    ElementTransformUtils.MoveElement(
                        destDoc,
                        destVportId,
                        new XYZ(dx, dy, 0));
                    t.Commit();
                }
            }
            catch (Exception ex)
            {
                result.Warnings.Add(
                    $"BBox position correction failed: "
                    + ex.Message);
            }
        }

        /// <summary>
        /// Gets the sheet number where a view is currently placed,
        /// or null if not placed.
        /// </summary>
        private static string GetViewSheetNumber(
            Document doc, View view)
        {
            try
            {
                var param = view.get_Parameter(
                    BuiltInParameter.VIEW_SHEET_VIEWPORT_INFO);
                if (param != null)
                {
                    string val = param.AsString();
                    if (!string.IsNullOrEmpty(val)
                    && val != "---"
                    && !val.StartsWith("Not ", StringComparison.OrdinalIgnoreCase))
                        return val;
                }
            }
            catch { }
            return null;
        }


        // ═══════════════════════════════════════════════════════
        //  7e. SHEET GUIDE GRIDS
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Copies and assigns guide grids from source sheet to
        /// destination sheet. Mirrors pyRevit copy_sheet_guides.
        /// </summary>
        private static void CopySheetGuides(
            Document source, ViewSheet srcSheet,
            Document target, ViewSheet destSheet,
            TransferResult result)
        {
            try
            {
                var guideParam = srcSheet.get_Parameter(
                    BuiltInParameter.SHEET_GUIDE_GRID);
                if (guideParam == null) return;

                Element srcGuide = source.GetElement(
                    guideParam.AsElementId());
                if (srcGuide == null) return;

                // Find or copy guide in target
                Element destGuide = FindGuideByName(
                    target, srcGuide.Name);

                if (destGuide == null)
                {
                    // Copy guide to target
                    var cpOpts = new CopyPasteOptions();
                    cpOpts.SetDuplicateTypeNamesHandler(
                        new UseDestinationTypesHandler());

                    using (var t = new Transaction(target,
                        "HMV – Copy Guide Grid"))
                    {
                        t.Start();
                        ElementTransformUtils.CopyElements(
                            source,
                            new List<ElementId> { srcGuide.Id },
                            target,
                            null,
                            cpOpts);
                        t.Commit();
                    }

                    destGuide = FindGuideByName(target, srcGuide.Name);
                }

                if (destGuide != null)
                {
                    using (var t = new Transaction(target,
                        "HMV – Set Sheet Guide Grid"))
                    {
                        t.Start();
                        var destParam = destSheet.get_Parameter(
                            BuiltInParameter.SHEET_GUIDE_GRID);
                        if (destParam != null)
                            destParam.Set(destGuide.Id);
                        t.Commit();
                    }
                }
                else
                {
                    result.Warnings.Add(
                        $"Could not copy guide grid for sheet "
                        + $"'{srcSheet.SheetNumber}'.");
                }
            }
            catch (Exception ex)
            {
                result.Warnings.Add(
                    $"Guide grid copy failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Finds a guide grid element by name in a document.
        /// Mirrors pyRevit find_guide.
        /// </summary>
        private static Element FindGuideByName(
            Document doc, string guideName)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_GuideGrid)
                .WhereElementIsNotElementType()
                .ToElements()
                .FirstOrDefault(e => string.Equals(
                    e.Name, guideName,
                    StringComparison.OrdinalIgnoreCase));
        }


        // ═══════════════════════════════════════════════════════
        //  7f. SHEET REVISIONS
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Copies revisions from source sheet to destination sheet.
        /// Creates missing revisions in target if needed.
        /// Mirrors pyRevit copy_sheet_revisions.
        /// </summary>
        private static void CopySheetRevisions(
            Document source, ViewSheet srcSheet,
            Document target, ViewSheet destSheet,
            TransferResult result)
        {
            try
            {
                var srcRevIds = srcSheet.GetAdditionalRevisionIds();
                if (srcRevIds == null || srcRevIds.Count == 0) return;

                var allDestRevs = new FilteredElementCollector(target)
                    .OfCategory(BuiltInCategory.OST_Revisions)
                    .WhereElementIsNotElementType()
                    .Cast<Revision>()
                    .ToList();

                var revisionsToSet = new List<ElementId>();

                using (var t = new Transaction(target,
                    "HMV – Copy Sheet Revisions"))
                {
                    t.Start();

                    foreach (var srcRevId in srcRevIds)
                    {
                        Revision srcRev = source.GetElement(srcRevId)
                            as Revision;
                        if (srcRev == null) continue;

                        ElementId destRevId = FindOrCreateRevision(
                            srcRev, allDestRevs, target, result);

                        if (destRevId != ElementId.InvalidElementId)
                            revisionsToSet.Add(destRevId);
                    }

                    if (revisionsToSet.Count > 0)
                    {
                        // Merge with existing additional revisions
                        var existing = destSheet
                            .GetAdditionalRevisionIds()
                            .ToList();
                        foreach (var id in revisionsToSet)
                        {
                            if (!existing.Contains(id))
                                existing.Add(id);
                        }
                        destSheet.SetAdditionalRevisionIds(existing);
                    }

                    t.Commit();
                }
            }
            catch (Exception ex)
            {
                result.Warnings.Add(
                    $"Revision copy failed for sheet "
                    + $"'{srcSheet.SheetNumber}': {ex.Message}");
            }
        }

        /// <summary>
        /// Finds a matching revision in target or creates a new one.
        /// Mirrors pyRevit ensure_dest_revision.
        /// </summary>
        private static ElementId FindOrCreateRevision(
            Revision srcRev,
            List<Revision> allDestRevs,
            Document target,
            TransferResult result)
        {
            // Try to match by date + description
            foreach (var destRev in allDestRevs)
            {
                if (destRev.RevisionDate == srcRev.RevisionDate
                    && destRev.Description == srcRev.Description)
                    return destRev.Id;
            }

            // Create new revision
            try
            {
                Revision newRev = Revision.Create(target);
                newRev.Description = srcRev.Description;
                newRev.IssuedBy = srcRev.IssuedBy;
                newRev.IssuedTo = srcRev.IssuedTo;
                newRev.RevisionDate = srcRev.RevisionDate;

                result.Warnings.Add(
                    $"Created revision: '{srcRev.Description}' "
                    + $"({srcRev.RevisionDate})");

                // Add to cache so we don't create it again
                allDestRevs.Add(newRev);

                return newRev.Id;
            }
            catch (Exception ex)
            {
                result.Warnings.Add(
                    $"Could not create revision "
                    + $"'{srcRev.Description}': {ex.Message}");
                return ElementId.InvalidElementId;
            }
        }


        // ═══════════════════════════════════════════════════════
        //  4. COPY & ASSIGN VIEW TEMPLATES
        // ═══════════════════════════════════════════════════════

        public static void CopyAndAssignViewTemplates(
            Document source,
            Document target,
            Dictionary<ElementId, ElementId> viewMap,
            TransferResult result)
        {
            if (viewMap.Count == 0) return;

            var opts = new CopyPasteOptions();
            opts.SetDuplicateTypeNamesHandler(
                new UseDestinationTypesHandler());

            var copiedTemplates = new Dictionary<ElementId, ElementId>();

            using (var t = new Transaction(target,
                "HMV – Copy & Assign View Templates"))
            {
                t.Start();

                foreach (var kvp in viewMap)
                {
                    View srcView = source.GetElement(kvp.Key) as View;
                    if (srcView == null) continue;
                    if (srcView.ViewTemplateId == ElementId.InvalidElementId)
                        continue;

                    ElementId srcTemplateId = srcView.ViewTemplateId;

                    if (!copiedTemplates.ContainsKey(srcTemplateId))
                    {
                        try
                        {
                            var copied = ElementTransformUtils.CopyElements(
                                source,
                                new List<ElementId> { srcTemplateId },
                                target,
                                Transform.Identity,
                                opts);

                            if (copied != null && copied.Count > 0)
                            {
                                copiedTemplates[srcTemplateId] =
                                    copied.First();
                            }
                        }
                        catch (Exception ex)
                        {
                            result.Warnings.Add(
                                $"Could not copy template "
                                + $"'{source.GetElement(srcTemplateId)?.Name}'"
                                + $": {ex.Message}");
                            continue;
                        }
                    }

                    if (copiedTemplates.TryGetValue(srcTemplateId,
                            out ElementId tgtTemplateId))
                    {
                        View tgtView = target.GetElement(kvp.Value) as View;
                        if (tgtView == null) continue;

                        try
                        {
                            tgtView.ViewTemplateId = tgtTemplateId;
                            result.TemplatesRemapped++;
                        }
                        catch
                        {
                            result.Warnings.Add(
                                $"Could not assign template to "
                                + $"'{tgtView.Name}'.");
                        }
                    }
                }

                t.Commit();
            }
        }

        // ═══════════════════════════════════════════════════════
        //  4b. COPY VIEW ANNOTATIONS (2ND PASS)
        // ═══════════════════════════════════════════════════════

        public static void CopyViewAnnotations(
            Document source,
            Document target,
            Dictionary<ElementId, ElementId> viewMap,
            TransferResult result)
        {
            if (viewMap.Count == 0) return;

            var opts = new CopyPasteOptions();
            opts.SetDuplicateTypeNamesHandler(
                new UseDestinationTypesHandler());

            var categories = new[]
            {
                BuiltInCategory.OST_Lines,
                BuiltInCategory.OST_TextNotes,
                BuiltInCategory.OST_GenericAnnotation,
                BuiltInCategory.OST_DetailComponents,
                BuiltInCategory.OST_FilledRegion,
                BuiltInCategory.OST_Dimensions,
                BuiltInCategory.OST_SpotElevations,
            };

            foreach (var kvp in viewMap)
            {
                View srcView = source.GetElement(kvp.Key) as View;
                View tgtView = target.GetElement(kvp.Value) as View;
                if (srcView == null || tgtView == null) continue;

                if (srcView.ViewType == ViewType.DraftingView
                    || srcView.ViewType == ViewType.Legend)
                    continue;

                foreach (var bic in categories)
                {
                    var ids = new FilteredElementCollector(
                            source, srcView.Id)
                        .OfCategory(bic)
                        .WhereElementIsNotElementType()
                        .ToElementIds()
                        .ToList();

                    if (ids.Count == 0) continue;

                    TryCopyAnnotations(
                        srcView, tgtView, ids, opts, result,
                        $"{bic} in '{srcView.Name}'");
                }
            }
        }

        private static bool TryCopyAnnotations(
            View srcView, View tgtView,
            ICollection<ElementId> ids,
            CopyPasteOptions opts,
            TransferResult result,
            string label)
        {
            Document target = tgtView.Document;
            using (var t = new Transaction(target,
                $"HMV – Copy {label}"))
            {
                t.Start();
                try
                {
                    var copied = ElementTransformUtils.CopyElements(
                        srcView, ids,
                        tgtView, Transform.Identity, opts);

                    result.AnnotationsCopied += copied.Count;
                    t.Commit();
                    return true;
                }
                catch
                {
                    t.RollBack();
                    result.Warnings.Add(
                        $"{label}: view-to-view copy failed.");
                    return false;
                }
            }
        }


        // ═══════════════════════════════════════════════════════
        //  6. COPY CATEGORY GRAPHIC OVERRIDES
        // ═══════════════════════════════════════════════════════

        public static void CopyCategoryOverrides(
            Document source,
            Document target,
            Dictionary<ElementId, ElementId> viewMap,
            TransferResult result)
        {
            if (viewMap.Count == 0) return;

            var fillPatternMap = new Dictionary<ElementId, ElementId>();
            var linePatternMap = new Dictionary<ElementId, ElementId>();

            var copyOpts = new CopyPasteOptions();
            copyOpts.SetDuplicateTypeNamesHandler(
                new UseDestinationTypesHandler());

            int overrideCount = 0;

            using (var t = new Transaction(target,
                "HMV – Copy Category Overrides"))
            {
                t.Start();

                foreach (var kvp in viewMap)
                {
                    View srcView = source.GetElement(kvp.Key) as View;
                    View tgtView = target.GetElement(kvp.Value) as View;
                    if (srcView == null || tgtView == null) continue;

                    if (srcView.ViewType == ViewType.DraftingView
                        || srcView.ViewType == ViewType.Legend)
                        continue;

                    foreach (Category cat in source.Settings.Categories)
                    {
                        if (cat == null) continue;
                        ElementId catId = cat.Id;

                        try
                        {
                            OverrideGraphicSettings srcOgs =
                                srcView.GetCategoryOverrides(catId);

                            if (IsOverrideEmpty(srcOgs)) continue;

                            OverrideGraphicSettings tgtOgs =
                                RemapOverride(source, target,
                                    srcOgs, fillPatternMap,
                                    linePatternMap, copyOpts);

                            tgtView.SetCategoryOverrides(catId, tgtOgs);
                            overrideCount++;
                        }
                        catch { }
                    }
                }

                t.Commit();
            }

            if (overrideCount > 0)
                result.Warnings.Add(
                    $"{overrideCount} category graphic overrides applied.");
        }

        private static OverrideGraphicSettings RemapOverride(
            Document source, Document target,
            OverrideGraphicSettings src,
            Dictionary<ElementId, ElementId> fillMap,
            Dictionary<ElementId, ElementId> lineMap,
            CopyPasteOptions copyOpts)
        {
            var tgt = new OverrideGraphicSettings();

            tgt.SetProjectionLineColor(src.ProjectionLineColor);
            tgt.SetCutLineColor(src.CutLineColor);
            tgt.SetSurfaceForegroundPatternColor(src.SurfaceForegroundPatternColor);
            tgt.SetSurfaceBackgroundPatternColor(src.SurfaceBackgroundPatternColor);
            tgt.SetCutForegroundPatternColor(src.CutForegroundPatternColor);
            tgt.SetCutBackgroundPatternColor(src.CutBackgroundPatternColor);

            tgt.SetProjectionLineWeight(src.ProjectionLineWeight);
            tgt.SetCutLineWeight(src.CutLineWeight);

            tgt.SetHalftone(src.Halftone);
            tgt.SetSurfaceTransparency(src.Transparency);

            tgt.SetProjectionLinePatternId(
                RemapLinePattern(source, target,
                    src.ProjectionLinePatternId,
                    lineMap, copyOpts));

            tgt.SetCutLinePatternId(
                RemapLinePattern(source, target,
                    src.CutLinePatternId,
                    lineMap, copyOpts));

            tgt.SetSurfaceForegroundPatternId(
                RemapFillPattern(source, target,
                    src.SurfaceForegroundPatternId,
                    fillMap, copyOpts));

            tgt.SetSurfaceBackgroundPatternId(
                RemapFillPattern(source, target,
                    src.SurfaceBackgroundPatternId,
                    fillMap, copyOpts));

            tgt.SetCutForegroundPatternId(
                RemapFillPattern(source, target,
                    src.CutForegroundPatternId,
                    fillMap, copyOpts));

            tgt.SetCutBackgroundPatternId(
                RemapFillPattern(source, target,
                    src.CutBackgroundPatternId,
                    fillMap, copyOpts));

            return tgt;
        }

        private static ElementId RemapFillPattern(
            Document source, Document target,
            ElementId srcPatternId,
            Dictionary<ElementId, ElementId> cache,
            CopyPasteOptions copyOpts)
        {
            if (srcPatternId == null
                || srcPatternId == ElementId.InvalidElementId)
                return ElementId.InvalidElementId;

            if (cache.TryGetValue(srcPatternId, out ElementId cached))
                return cached;

            FillPatternElement srcElem =
                source.GetElement(srcPatternId) as FillPatternElement;
            if (srcElem == null)
            {
                cache[srcPatternId] = ElementId.InvalidElementId;
                return ElementId.InvalidElementId;
            }

            FillPatternElement tgtElem =
                FillPatternElement.GetFillPatternElementByName(
                    target, srcElem.GetFillPattern().Target,
                    srcElem.Name);

            if (tgtElem != null)
            {
                cache[srcPatternId] = tgtElem.Id;
                return tgtElem.Id;
            }

            try
            {
                var copied = ElementTransformUtils.CopyElements(
                    source,
                    new List<ElementId> { srcPatternId },
                    target, Transform.Identity, copyOpts);

                ElementId newId = copied.FirstOrDefault()
                    ?? ElementId.InvalidElementId;
                cache[srcPatternId] = newId;
                return newId;
            }
            catch
            {
                cache[srcPatternId] = ElementId.InvalidElementId;
                return ElementId.InvalidElementId;
            }
        }

        private static ElementId RemapLinePattern(
            Document source, Document target,
            ElementId srcPatternId,
            Dictionary<ElementId, ElementId> cache,
            CopyPasteOptions copyOpts)
        {
            if (srcPatternId == null
                || srcPatternId == ElementId.InvalidElementId)
                return ElementId.InvalidElementId;

            if (srcPatternId.IntegerValue < 0)
                return srcPatternId;

            if (cache.TryGetValue(srcPatternId, out ElementId cached))
                return cached;

            LinePatternElement srcElem =
                source.GetElement(srcPatternId) as LinePatternElement;
            if (srcElem == null)
            {
                cache[srcPatternId] = ElementId.InvalidElementId;
                return ElementId.InvalidElementId;
            }

            LinePatternElement tgtElem =
                LinePatternElement.GetLinePatternElementByName(
                    target, srcElem.Name);

            if (tgtElem != null)
            {
                cache[srcPatternId] = tgtElem.Id;
                return tgtElem.Id;
            }

            try
            {
                var copied = ElementTransformUtils.CopyElements(
                    source,
                    new List<ElementId> { srcPatternId },
                    target, Transform.Identity, copyOpts);

                ElementId newId = copied.FirstOrDefault()
                    ?? ElementId.InvalidElementId;
                cache[srcPatternId] = newId;
                return newId;
            }
            catch
            {
                cache[srcPatternId] = ElementId.InvalidElementId;
                return ElementId.InvalidElementId;
            }
        }

        private static bool IsOverrideEmpty(OverrideGraphicSettings ogs)
        {
            var def = new OverrideGraphicSettings();

            return ColorsEqual(ogs.ProjectionLineColor, def.ProjectionLineColor)
                && ogs.ProjectionLineWeight == def.ProjectionLineWeight
                && ogs.ProjectionLinePatternId == def.ProjectionLinePatternId
                && ColorsEqual(ogs.CutLineColor, def.CutLineColor)
                && ogs.CutLineWeight == def.CutLineWeight
                && ogs.CutLinePatternId == def.CutLinePatternId
                && ColorsEqual(ogs.SurfaceForegroundPatternColor, def.SurfaceForegroundPatternColor)
                && ColorsEqual(ogs.SurfaceBackgroundPatternColor, def.SurfaceBackgroundPatternColor)
                && ogs.Transparency == def.Transparency
                && ogs.Halftone == def.Halftone;
        }

        private static bool ColorsEqual(Color a, Color b)
        {
            if (a == null && b == null) return true;
            if (a == null || b == null) return false;
            return a.Red == b.Red
                && a.Green == b.Green
                && a.Blue == b.Blue;
        }

        // ═══════════════════════════════════════════════════════
        //  5. REFERENCE MARKER IDENTIFICATION
        // ═══════════════════════════════════════════════════════

        public static List<string> IdentifyReferenceMarkers(
            Document source, View sourceView)
        {
            var markers = new List<string>();

            CollectRefs(source, sourceView, "Section", markers,
                v => { var r = v.GetReferenceSections(); return r != null ? r.ToList() : new List<ElementId>(); });
            CollectRefs(source, sourceView, "Callout", markers,
                v => { var r = v.GetReferenceCallouts(); return r != null ? r.ToList() : new List<ElementId>(); });
            CollectRefs(source, sourceView, "Elevation", markers,
                v => { var r = v.GetReferenceElevations(); return r != null ? r.ToList() : new List<ElementId>(); });

            return markers;
        }

        private static void CollectRefs(
            Document doc,
            View view,
            string label,
            List<string> markers,
            Func<View, List<ElementId>> getter)
        {
            try
            {
                List<ElementId> ids = getter(view);
                foreach (var id in ids)
                {
                    Element el = doc.GetElement(id);
                    if (el != null)
                        markers.Add(
                            $"{label}: {el.Name} "
                            + $"(Id {id.IntegerValue})");
                }
            }
            catch { /* not available on all view types */ }
        }


        // ═══════════════════════════════════════════════════════
        //  HELPERS: ForceCreate, Level matching, Unique names
        // ══════════════════════════════════════════════

        private static ElementId ForceCreateView(
            Document source, Document target,
            ElementId srcViewId, TransferResult result)
        {
            View srcView = source.GetElement(srcViewId) as View;
            if (srcView == null) return ElementId.InvalidElementId;

            try
            {
                if (srcView is ViewPlan srcPlan)
                    return ForceCreatePlan(source, target, srcPlan, result);

                if (srcView is ViewSection srcSection)
                    return ForceCreateSection(source, target, srcSection, result);

                return ElementId.InvalidElementId;
            }
            catch (Exception ex)
            {
                result.Warnings.Add(
                    $"ForceCreate failed for '{srcView.Name}': {ex.Message}");
                return ElementId.InvalidElementId;
            }
        }

        private static ElementId ForceCreatePlan(
            Document source, Document target,
            ViewPlan srcPlan, TransferResult result)
        {
            if (srcPlan == null || srcPlan.GenLevel == null)
                return ElementId.InvalidElementId;

            Level srcLevel = source.GetElement(
                srcPlan.GenLevel.Id) as Level;
            if (srcLevel == null) return ElementId.InvalidElementId;

            Level tgtLevel = FindClosestLevel(
                target, srcLevel.Elevation);
            if (tgtLevel == null) return ElementId.InvalidElementId;

            // Find existing ViewPlan on that level to duplicate
            ViewPlan existingView = new FilteredElementCollector(target)
                .OfClass(typeof(ViewPlan))
                .Cast<ViewPlan>()
                .FirstOrDefault(v => !v.IsTemplate
                    && v.GenLevel != null
                    && v.GenLevel.Id == tgtLevel.Id
                    && v.ViewType == srcPlan.ViewType);

            if (existingView == null)
            {
                result.Warnings.Add(
                    $"No existing view on level '{tgtLevel.Name}' "
                    + "to duplicate.");
                return ElementId.InvalidElementId;
            }

            ViewPlan newView = null;
            using (var t = new Transaction(target, "HMV – Create Plan"))
            {
                t.Start();
                try
                {
                    ElementId newId = existingView.Duplicate(
                        ViewDuplicateOption.WithDetailing);

                    newView = target.GetElement(newId) as ViewPlan;
                    if (newView == null)
                    {
                        t.RollBack();
                        result.Warnings.Add(
                            $"[DBG-FP] Duplicate failed for '{srcPlan.Name}'.");
                        return ElementId.InvalidElementId;
                    }

                    newView.Name = GetUniqueViewName(target, srcPlan.Name);

                    // ── Full CropBox copy (includes rotation transform) ──
                    try
                    {
                        newView.CropBox = srcPlan.CropBox;
                        newView.CropBoxActive = srcPlan.CropBoxActive;
                        newView.CropBoxVisible = srcPlan.CropBoxVisible;
                    }
                    catch { }

                    // ── View range ──
                    try
                    {
                        PlanViewRange srcRange = srcPlan.GetViewRange();
                        PlanViewRange tgtRange = newView.GetViewRange();
                        tgtRange.SetOffset(PlanViewPlane.TopClipPlane,
                            srcRange.GetOffset(PlanViewPlane.TopClipPlane));
                        tgtRange.SetOffset(PlanViewPlane.CutPlane,
                            srcRange.GetOffset(PlanViewPlane.CutPlane));
                        tgtRange.SetOffset(PlanViewPlane.BottomClipPlane,
                            srcRange.GetOffset(PlanViewPlane.BottomClipPlane));
                        tgtRange.SetOffset(PlanViewPlane.ViewDepthPlane,
                            srcRange.GetOffset(PlanViewPlane.ViewDepthPlane));
                        newView.SetViewRange(tgtRange);
                    }
                    catch { }

                    // ── Scale, DetailLevel, DisplayStyle, Description ──
                    CopyViewProperties(srcPlan, newView);

                    t.Commit();

                    result.Warnings.Add(
                        $"[DBG-FP] '{srcPlan.Name}': created plan OK "
                        + $"id={newView.Id.IntegerValue} on level '{tgtLevel.Name}'");
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    result.Warnings.Add(
                        $"[DBG-FP] '{srcPlan.Name}': FAILED: {ex.Message}");
                    return ElementId.InvalidElementId;
                }
            }

            return newView.Id;
        }

        private static ElementId ForceCreateSection(
            Document source, Document target,
            ViewSection srcSection, TransferResult result)
        {
            if (srcSection == null)
                return ElementId.InvalidElementId;

            if (srcSection.ViewType == ViewType.Elevation)
            {
                result.Warnings.Add(
                    $"Elevation '{srcSection.Name}' cannot be created "
                    + "programmatically. Use Update mode or place manually.");
                return ElementId.InvalidElementId;
            }

            // ── STEP 1: Find donor section to duplicate ───────────
            ViewSection donor = new FilteredElementCollector(target)
                .OfClass(typeof(ViewSection))
                .Cast<ViewSection>()
                .FirstOrDefault(v =>
                    !v.IsTemplate
                    && v.ViewType == srcSection.ViewType);

            if (donor == null)
            {
                donor = new FilteredElementCollector(target)
                    .OfClass(typeof(ViewSection))
                    .Cast<ViewSection>()
                    .FirstOrDefault(v => !v.IsTemplate);
            }

            if (donor == null)
            {
                result.Warnings.Add(
                    $"No existing section in target to duplicate "
                    + $"for '{srcSection.Name}'. "
                    + "Create at least one section manually in the target.");
                return ElementId.InvalidElementId;
            }
            // ── STEP 1b: Snapshot existing OST_Viewers before duplicate ──
            var existingViewerIds = new HashSet<int>(
                new FilteredElementCollector(target)
                    .OfCategory(BuiltInCategory.OST_Viewers)
                    .WhereElementIsNotElementType()
                    .ToElementIds()
                    .Select(id => id.IntegerValue));

            // ── STEP 2: Duplicate + set CropBox ──────────────────
            ViewSection newView = null;
            using (var t = new Transaction(target, "HMV – Duplicate Section"))
            {
                t.Start();
                try
                {
                    ElementId newId = donor.Duplicate(
                        ViewDuplicateOption.Duplicate);
                    newView = target.GetElement(newId) as ViewSection;

                    if (newView == null)
                    {
                        t.RollBack();
                        result.Warnings.Add(
                            $"Duplicate failed for '{srcSection.Name}'.");
                        return ElementId.InvalidElementId;
                    }

                    // ── DEBUG 1: source view model-space geometry ──
                    result.Warnings.Add(
                        $"[DBG-SEC1] '{srcSection.Name}': "
                        + $"src.Origin=({srcSection.Origin.X:F4},{srcSection.Origin.Y:F4},{srcSection.Origin.Z:F4}) "
                        + $"src.ViewDir=({srcSection.ViewDirection.X:F4},{srcSection.ViewDirection.Y:F4},{srcSection.ViewDirection.Z:F4}) "
                        + $"src.UpDir=({srcSection.UpDirection.X:F4},{srcSection.UpDirection.Y:F4},{srcSection.UpDirection.Z:F4}) "
                        + $"src.CropOrigin=({srcSection.CropBox.Transform.Origin.X:F4},"
                        + $"{srcSection.CropBox.Transform.Origin.Y:F4},"
                        + $"{srcSection.CropBox.Transform.Origin.Z:F4})");

                    // ── DEBUG 2: donor view BEFORE overwrite ──
                    result.Warnings.Add(
                        $"[DBG-SEC2] donor='{donor.Name}': "
                        + $"donor.Origin=({donor.Origin.X:F4},{donor.Origin.Y:F4},{donor.Origin.Z:F4}) "
                        + $"donor.CropOrigin=({donor.CropBox.Transform.Origin.X:F4},"
                        + $"{donor.CropBox.Transform.Origin.Y:F4},"
                        + $"{donor.CropBox.Transform.Origin.Z:F4})");

                    newView.Name = GetUniqueViewName(target, srcSection.Name);

                    newView.CropBox        = srcSection.CropBox;
                    newView.CropBoxActive  = srcSection.CropBoxActive;
                    newView.CropBoxVisible = srcSection.CropBoxVisible;
                    // ── DEBUG 3: newView AFTER CropBox overwrite, BEFORE commit ──
                    result.Warnings.Add(
                        $"[DBG-SEC3] '{srcSection.Name}': "
                        + $"new.Origin=({newView.Origin.X:F4},{newView.Origin.Y:F4},{newView.Origin.Z:F4}) "
                        + $"new.ViewDir=({newView.ViewDirection.X:F4},{newView.ViewDirection.Y:F4},{newView.ViewDirection.Z:F4}) "
                        + $"new.CropOrigin=({newView.CropBox.Transform.Origin.X:F4},"
                        + $"{newView.CropBox.Transform.Origin.Y:F4},"
                        + $"{newView.CropBox.Transform.Origin.Z:F4})");

                    t.Commit();
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    result.Warnings.Add(
                        $"Section duplicate/cropbox failed for "
                        + $"'{srcSection.Name}': {ex.Message}");
                    return ElementId.InvalidElementId;
                }
            }

            // ── STEP 3: Compute delta ─────────────────────────────
            // ── DEBUG 4: newView AFTER commit ──
            result.Warnings.Add(
                $"[DBG-SEC4] '{srcSection.Name}': "
                + $"postCommit.Origin=({newView.Origin.X:F4},{newView.Origin.Y:F4},{newView.Origin.Z:F4}) "
                + $"postCommit.ViewDir=({newView.ViewDirection.X:F4},{newView.ViewDirection.Y:F4},{newView.ViewDirection.Z:F4}) "
                + $"postCommit.CropOrigin=({newView.CropBox.Transform.Origin.X:F4},"
                + $"{newView.CropBox.Transform.Origin.Y:F4},"
                + $"{newView.CropBox.Transform.Origin.Z:F4})");
            XYZ srcOrigin  = srcSection.CropBox.Transform.Origin;
            XYZ destOrigin = newView.CropBox.Transform.Origin;
            XYZ delta      = srcOrigin - destOrigin;

            result.Warnings.Add(
                $"[DEBUG-SEC] '{srcSection.Name}': "
                + $"delta=({delta.X:F4},{delta.Y:F4},{delta.Z:F4})");

            // ── STEP 4: Find NEW viewer by snapshot diff ──────────
            bool needsMove = delta.GetLength() > 1e-4;

            if (needsMove)
            {
                Element viewer = new FilteredElementCollector(target)
                    .OfCategory(BuiltInCategory.OST_Viewers)
                    .WhereElementIsNotElementType()
                    .ToElements()
                    .FirstOrDefault(e =>
                        !existingViewerIds.Contains(e.Id.IntegerValue));

                if (viewer != null)
                {
                    // ── DEBUG 5: viewer position BEFORE move ──
                    BoundingBoxXYZ vBB = viewer.get_BoundingBox(null);
                    XYZ vCenter = (vBB != null)
                        ? (vBB.Min + vBB.Max) / 2.0
                        : XYZ.Zero;
                    result.Warnings.Add(
                        $"[DBG-SEC5] '{srcSection.Name}': "
                        + $"viewer.Id={viewer.Id.IntegerValue} "
                        + $"viewer.Category={viewer.Category?.Name ?? "null"} "
                        + $"viewerCenter=({vCenter.X:F4},{vCenter.Y:F4},{vCenter.Z:F4}) "
                        + $"delta=({delta.X:F4},{delta.Y:F4},{delta.Z:F4}) len={delta.GetLength():F6}");
                    using (var t = new Transaction(target,
                        "HMV – Move + Orient Section"))
                    {
                        t.Start();
                        try
                        {
                            // ── MOVE ──
                            ElementTransformUtils.MoveElement(
                                target, viewer.Id, delta);

                            // ── ROTATE to match source ViewDirection ──
                            XYZ srcDir = srcSection.ViewDirection;
                            XYZ newDir = newView.ViewDirection;

                            // Project both onto XY plane (ignore Z for vertical sections)
                            XYZ srcDirXY = new XYZ(srcDir.X, srcDir.Y, 0).Normalize();
                            XYZ newDirXY = new XYZ(newDir.X, newDir.Y, 0).Normalize();

                            double dot = srcDirXY.DotProduct(newDirXY);
                            double cross = srcDirXY.X * newDirXY.Y
                                          - srcDirXY.Y * newDirXY.X;
                            double angle = Math.Atan2(cross, dot);

                            if (Math.Abs(angle) > 1e-6)
                            {
                                // Pivot = source origin (where the section should end up)
                                XYZ pivot = srcSection.Origin;
                                Line axis = Line.CreateBound(
                                    pivot,
                                    pivot + XYZ.BasisZ);

                                // Move first placed the viewer near srcOrigin,
                                // now rotate around that point
                                ElementTransformUtils.RotateElement(
                                    target, viewer.Id, axis, -angle);

                                result.Warnings.Add(
                                    $"[DBG-ROT] '{srcSection.Name}': "
                                    + $"rotated {angle * 180.0 / Math.PI:F2}° "
                                    + $"srcDir=({srcDir.X:F4},{srcDir.Y:F4},{srcDir.Z:F4}) "
                                    + $"donorDir=({newDir.X:F4},{newDir.Y:F4},{newDir.Z:F4})");
                            }

                            t.Commit();

                            // ── DEBUG: final position check ──
                            result.Warnings.Add(
                                $"[DBG-FINAL] '{srcSection.Name}': "
                                + $"final.Origin=({newView.Origin.X:F4},{newView.Origin.Y:F4},{newView.Origin.Z:F4}) "
                                + $"final.ViewDir=({newView.ViewDirection.X:F4},{newView.ViewDirection.Y:F4},{newView.ViewDirection.Z:F4}) "
                                + $"expected.Origin=({srcSection.Origin.X:F4},{srcSection.Origin.Y:F4},{srcSection.Origin.Z:F4})");
                        }
                        catch (Exception ex)
                        {
                            t.RollBack();
                            result.Warnings.Add(
                                $"[DBG-SEC] '{srcSection.Name}': "
                                + $"move/rotate failed: {ex.Message}");
                        }
                    }
                }
                else
                {
                    result.Warnings.Add(
                        $"[DEBUG-SEC] '{srcSection.Name}': "
                        + "new OST_Viewers element not found.");
                }
            }

            // ── STEP 5: Copy scale/properties (NO CropBox here) ──
            // CopyViewProperties sets CropBox again — we skip that
            // by calling only the individual properties we need.
            try { newView.Scale       = srcSection.Scale;       } catch { }
            try { newView.DetailLevel = srcSection.DetailLevel; } catch { }
            try { newView.DisplayStyle= srcSection.DisplayStyle;} catch { }
            try
            {
                var sp = srcSection.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION);
                var dp = newView.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION);
                if (sp != null && dp != null && !dp.IsReadOnly)
                    dp.Set(sp.AsString() ?? "");
            }
            catch { }

            // ── STEP 6: Copy view contents (annotations, dims, etc.) ──
            using (var t = new Transaction(target, "HMV – Copy Section Contents"))
            {
                SetSwallowErrors(t);
                t.Start();
                try
                {
                    var contentIds = GetViewContents(
                        source, srcSection, new SheetCopyOptions());

                    if (contentIds.Count > 0)
                    {
                        var cpOpts = new CopyPasteOptions();
                        cpOpts.SetDuplicateTypeNamesHandler(
                            new UseDestinationTypesHandler());

                        ElementTransformUtils.CopyElements(
                            srcSection,
                            contentIds,
                            newView,
                            null,
                            cpOpts);

                        result.Warnings.Add(
                            $"[DEBUG-SEC] '{srcSection.Name}': "
                            + $"copied {contentIds.Count} content elements.");
                    }
                }
                catch
                {
                    // Swallow — commit whatever succeeded
                }

                if (t.HasStarted() && !t.HasEnded())
                {
                    try { t.Commit(); }
                    catch { try { t.RollBack(); } catch { } }
                }
            }

            return newView.Id;
        }

        private static Level FindClosestLevel(Document doc, double elevation)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => Math.Abs(l.Elevation - elevation))
                .FirstOrDefault();
        }

        private static string GetUniqueViewName(Document doc, string desired)
        {
            var existing = new HashSet<string>(
                new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => !v.IsTemplate)
                    .Select(v => v.Name));

            if (!existing.Contains(desired)) return desired;

            int n = 2;
            while (existing.Contains($"{desired} ({n})")) n++;
            return $"{desired} ({n})";
        }
    }


    // ═══════════════════════════════════════════════════════════
    //  ViewportData — captures source viewport placement info
    // ═══════════════════════════════════════════════════════════

    /// <summary>
    /// Holds source viewport positioning data for accurate
    /// placement in the destination sheet.
    /// </summary>
    public class ViewportData
    {
        public XYZ BoxCenter { get; set; }
        public XYZ LabelOffset { get; set; }
        public double? LabelLineLength { get; set; }
        public XYZ BBoxMin { get; set; }
        public XYZ BBoxMax { get; set; }
        public ViewportRotation Rotation { get; set; } = ViewportRotation.None;
    }
    public class SwallowErrorsHandler : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(
            FailuresAccessor failuresAccessor)
        {
            var failures = failuresAccessor.GetFailureMessages();

            foreach (var f in failures)
            {
                // Delete warnings (non-critical)
                if (f.GetSeverity() == FailureSeverity.Warning)
                {
                    failuresAccessor.DeleteWarning(f);
                }
                // Try to resolve errors
                else if (f.HasResolutions())
                {
                    failuresAccessor.ResolveFailure(f);
                }
                else
                {
                    failuresAccessor.DeleteWarning(f);
                }
            }

            return FailureProcessingResult.Continue;
        }
    }
}