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
            TransferResult result)
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
                CopyViewBatch(source, target,
                    flatIds, Transform.Identity,
                    opts, viewMap, result,
                    "Drafting/Legend", mode);
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
            using (var t = new Transaction(target,
                $"HMV – Copy {batchLabel} Views"))
            {
                t.Start();

                var existingViewIds = new HashSet<int>(
                    new FilteredElementCollector(target)
                        .OfClass(typeof(View))
                        .Select(v => v.Id.IntegerValue));
                int batchSuccesses = 0;
                int failures = 0;

                foreach (var srcId in ids)
                {
                    string viewName = "(unknown)";
                    View srcV = null;
                    try
                    {
                        srcV = source.GetElement(srcId) as View;
                        if (srcV != null) viewName = srcV.Name;
                    }
                    catch { }

                    // ── UPDATE MODE: find existing, clear, recopy ──
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
                        // If no match found in Update mode, fall through
                        // to create a new view below.
                    }

                    // ── CREATE MODE (or no match in Update) ───────
                    bool needsForcePath = false;

                    if (srcV is ViewPlan srcPlan && srcPlan.GenLevel != null)
                    {
                        Level srcLev = source.GetElement(srcPlan.GenLevel.Id) as Level;
                        if (srcLev != null)
                        {
                            Level tgtLev = FindClosestLevel(target, srcLev.Elevation);
                            if (tgtLev != null
                                && Math.Abs(tgtLev.Elevation - srcLev.Elevation) < 0.01)
                            {
                                needsForcePath = true;
                            }
                        }
                    }
                    else if (srcV is ViewSection)
                    {
                        needsForcePath = true;
                    }

                    using (var sub = new SubTransaction(target))
                    {
                        try
                        {
                            sub.Start();

                            if (needsForcePath)
                            {
                                ElementId forced = ForceCreateView(
                                    source, target, srcId, result);

                                if (forced != null
                                    && forced != ElementId.InvalidElementId)
                                {
                                    viewMap[srcId] = forced;
                                    result.ViewsCreated++;
                                    batchSuccesses++;
                                    sub.Commit();
                                }
                                else
                                {
                                    sub.RollBack();
                                    failures++;
                                    result.Warnings.Add(
                                        $"View '{viewName}': could not "
                                        + "duplicate in target.");
                                }
                            }
                            else
                            {
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

                if (viewMap.Count > 0 || failures < ids.Count)
                {
                    t.Commit();
                }
                else
                {
                    t.RollBack();
                    result.Errors.Add(
                        $"Copy {batchLabel} views: all {ids.Count} views failed.");
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
            try { destView.Scale = sourceView.Scale; } catch { }

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

        /// <summary>
        /// Copies or updates a single sheet.
        /// Mirrors pyRevit copy_sheet.
        /// </summary>
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

            if (destSheet != null)
            {
                if (mode == ViewTransferMode.Update)
                {
                    // Update existing: clear and recopy contents
                    result.Warnings.Add(
                        $"Sheet '{srcSheet.SheetNumber}' exists "
                        + "in target — updating contents.");

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
                    // Create mode but sheet exists — skip content
                    // copy, just report and proceed to viewports
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
                        destSheet = ViewSheet.CreatePlaceholder(target);
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

                // Copy sheet-level content (title block, schedules)
                using (var t = new Transaction(target,
                    "HMV – Copy Sheet Contents"))
                {
                    t.Start();
                    CopyViewContentsViewToView(
                        source, srcSheet, target, destSheet, opts);
                    t.Commit();
                }
            }

            if (destSheet == null) return null;

            // ── Viewports ─────────────────────────────────────
            if (!destSheet.IsPlaceholder && opts.CopyViewports)
            {
                CopySheetViewports(
                    source, srcSheet,
                    target, destSheet,
                    coordTransform, viewMap,
                    result, settings);
            }

            // ── Guide grids ──────────────────────────────────
            if (!destSheet.IsPlaceholder && opts.CopyGuideGrids)
            {
                CopySheetGuides(
                    source, srcSheet, target, destSheet, result);
            }

            // ── Revisions ────────────────────────────────────
            if (opts.CopyRevisions)
            {
                CopySheetRevisions(
                    source, srcSheet, target, destSheet, result);
            }

            return destSheet;
        }


        // ═══════════════════════════════════════════════════════
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

                View srcView = source.GetElement(srcVport.ViewId) as View;
                if (srcView == null) continue;

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

                if (destViewId == ElementId.InvalidElementId)
                {
                    result.Warnings.Add(
                        $"Could not copy view '{srcView.Name}' "
                        + "for viewport placement.");
                    continue;
                }

                // Check if view is already on another sheet
                View destView = target.GetElement(destViewId) as View;
                if (destView == null) continue;

                // Check existing placement
                if (existingViewIds.Contains(destViewId))
                {
                    result.Warnings.Add(
                        $"View '{srcView.Name}' already on sheet "
                        + $"'{destSheet.SheetNumber}'.");
                    continue;
                }

                // Check if placed on a different sheet
                string placedSheet = GetViewSheetNumber(target, destView);
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

                // ── Place viewport on destination sheet ───────
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

                        // Preserve detail number
                        if (settings.SheetOptions.PreserveDetailNumbers)
                        {
                            ApplyDetailNumber(
                                srcVport, newVport, result);
                        }

                        t.Commit();
                        result.ViewportsCopied++;
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

                // ── BBox position correction ──────────────────
                CorrectViewportByBBox(
                    target, newVport.Id, srcData, destSheet, result);
            }
        }


        // ═══════════════════════════════════════════════════════
        //  7b. SINGLE VIEW COPY/UPDATE FOR VIEWPORT REFS
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Copies or updates a single view referenced by a viewport.
        /// Uses the same create/update logic but tailored for the
        /// sheet-viewport workflow. Handles DrawingSheet, DraftingView,
        /// Legend, and spatial views.
        /// </summary>
        private static ElementId CopyOrUpdateSingleView(
            Document source, Document target,
            View srcView,
            Transform coordTransform,
            TransferResult result,
            MigrationSettings settings)
        {
            ViewTransferMode mode = settings.TransferMode;
            var opts = settings.SheetOptions;

            // Check for existing match
            View existing = FindMatchingView(target, srcView);

            if (existing != null && mode == ViewTransferMode.Update)
            {
                // Update existing view contents
                using (var t = new Transaction(target,
                    "HMV – Update View Contents"))
                {
                    t.Start();
                    ClearViewContents(target, existing, opts);
                    CopyViewContentsViewToView(
                        source, srcView, target, existing, opts);
                    CopyViewProperties(srcView, existing);
                    t.Commit();
                }
                result.ViewsUpdated++;
                return existing.Id;
            }
            else if (existing != null && mode == ViewTransferMode.Create)
            {
                // View exists but user wants Create mode
                // For views that can't be on multiple sheets, just
                // reuse the existing one
                result.Warnings.Add(
                    $"View '{srcView.Name}' exists in target, reusing.");
                return existing.Id;
            }

            // ── Create new view ───────────────────────────────
            ElementId newViewId = ElementId.InvalidElementId;

            using (var t = new Transaction(target,
                "HMV – Create View for Sheet"))
            {
                t.Start();

                switch (srcView.ViewType)
                {
                    case ViewType.DraftingView:
                        newViewId = CreateDraftingView(
                            source, target, srcView, result);
                        break;

                    case ViewType.Legend:
                        newViewId = CreateLegendView(
                            source, target, srcView, result);
                        break;

                    case ViewType.FloorPlan:
                    case ViewType.CeilingPlan:
                        newViewId = ForceCreatePlan(
                            source, target,
                            srcView as ViewPlan, result);
                        break;

                    case ViewType.Section:
                    case ViewType.Detail:
                    case ViewType.Elevation:
                        newViewId = ForceCreateSection(
                            source, target,
                            srcView as ViewSection, result);
                        break;

                    default:
                        result.Warnings.Add(
                            $"Unsupported view type "
                            + $"'{srcView.ViewType}' for "
                            + $"'{srcView.Name}'.");
                        t.RollBack();
                        return ElementId.InvalidElementId;
                }

                if (newViewId != ElementId.InvalidElementId)
                {
                    t.Commit();
                    result.ViewsCreated++;
                }
                else
                {
                    t.RollBack();
                    return ElementId.InvalidElementId;
                }
            }

            // Copy contents into the new view
            View newView = target.GetElement(newViewId) as View;
            if (newView != null)
            {
                using (var t = new Transaction(target,
                    "HMV – Copy View Contents"))
                {
                    t.Start();
                    CopyViewContentsViewToView(
                        source, srcView, target, newView, opts);
                    t.Commit();
                }
            }

            return newViewId;
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

            try { data.LabelOffset = vport.LabelOffset; }
            catch { /* Revit 2022+ only */ }

            try { data.LabelLineLength = vport.LabelLineLength; }
            catch { /* Revit 2022+ only */ }

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
                        && val != "---")
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
        // ═══════════════════════════════════════════════════════

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
            Level srcLevel = source.GetElement(srcPlan.GenLevel.Id) as Level;
            if (srcLevel == null) return ElementId.InvalidElementId;

            Level tgtLevel = FindClosestLevel(target, srcLevel.Elevation);
            if (tgtLevel == null) return ElementId.InvalidElementId;

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

            ElementId newId = existingView.Duplicate(
                ViewDuplicateOption.WithDetailing);

            ViewPlan newView = target.GetElement(newId) as ViewPlan;
            if (newView == null) return ElementId.InvalidElementId;

            newView.Name = GetUniqueViewName(target, srcPlan.Name);

            if (srcPlan.CropBoxActive)
            {
                try
                {
                    newView.CropBox = srcPlan.CropBox;
                    newView.CropBoxActive = true;
                    newView.CropBoxVisible = srcPlan.CropBoxVisible;
                }
                catch { }
            }

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

            try { newView.Scale = srcPlan.Scale; } catch { }

            return newView.Id;
        }

        private static ElementId ForceCreateSection(
            Document source, Document target,
            ViewSection srcSection, TransferResult result)
        {
            ViewFamily vf = srcSection.ViewType == ViewType.Detail
                ? ViewFamily.Detail : ViewFamily.Section;

            ElementId vftId = new FilteredElementCollector(target)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .Where(x => x.ViewFamily == vf)
                .Select(x => x.Id)
                .FirstOrDefault();

            if (vftId == null || vftId == ElementId.InvalidElementId)
                return ElementId.InvalidElementId;

            Transform coordTransform =
                ComputeSharedCoordinateTransform(source, target);

            BoundingBoxXYZ srcBox = srcSection.CropBox;
            Transform srcT = srcBox.Transform;

            BoundingBoxXYZ tgtBox = new BoundingBoxXYZ();

            Transform newT = Transform.Identity;
            newT.Origin = coordTransform.OfPoint(srcT.Origin);
            newT.BasisX = coordTransform.OfVector(srcT.BasisX).Normalize();
            newT.BasisY = coordTransform.OfVector(srcT.BasisY).Normalize();
            newT.BasisZ = coordTransform.OfVector(srcT.BasisZ).Normalize();

            tgtBox.Transform = newT;
            tgtBox.Min = srcBox.Min;
            tgtBox.Max = srcBox.Max;
            tgtBox.Enabled = true;

            ViewSection newView = ViewSection.CreateSection(
                target, vftId, tgtBox);

            newView.Name = GetUniqueViewName(target, srcSection.Name);
            newView.CropBoxActive = srcSection.CropBoxActive;
            newView.CropBoxVisible = srcSection.CropBoxVisible;
            try { newView.Scale = srcSection.Scale; } catch { }

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
    }
}