using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;

namespace HMVTools
{
    // ── ViewFactory ────────────────────────────────────────────
    //  View creation, duplication, batch copying, and matching.
    //  Mirrors pyRevit copy_view / force_create_* patterns.
    // ──────────────────────────────────────────────────────────

    public static class ViewFactory
    {
        // ══════════════════════════════════════════════════════
        //  VIEW MATCHING
        // ══════════════════════════════════════════════════════

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


        // ══════════════════════════════════════════════════════
        //  VIEW BATCH COPY
        // ══════════════════════════════════════════════════════

        internal static void CopyViewBatch(
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
            var forcePathIds = new List<ElementId>();
            var regularIds   = new List<ElementId>();

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

            // ══════════════════════════════════════════════════
            //  PASS 1 – regular views (inside outer Transaction)
            // ══════════════════════════════════════════════════
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
                                        ViewContentCopier.ClearViewContents(target, existing,
                                            new SheetCopyOptions());
                                        ViewContentCopier.CopyViewContentsViewToView(
                                            source, srcV, target, existing,
                                            new SheetCopyOptions());
                                        ViewContentCopier.CopyViewProperties(srcV, existing);
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

            // ══════════════════════════════════════════════════
            //  PASS 2 – force-path views (sections/details/elevations)
            //  Each ForceCreateSection opens its OWN transactions,
            //  so they MUST run outside any outer Transaction.
            // ══════════════════════════════════════════════════
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
                                ViewContentCopier.ClearViewContents(target, existing,
                                    new SheetCopyOptions());
                                ViewContentCopier.CopyViewContentsViewToView(
                                    source, srcV, target, existing,
                                    new SheetCopyOptions());
                                ViewContentCopier.CopyViewProperties(srcV, existing);
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


        // ══════════════════════════════════════════════════════
        //  COPY OR UPDATE SINGLE VIEW
        // ══════════════════════════════════════════════════════

        internal static ElementId CopyOrUpdateSingleView(
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
                        ViewContentCopier.ClearViewContents(target, matchingView, opts);
                        ViewContentCopier.CopyViewContentsViewToView(
                            source, srcView, target, matchingView, opts);
                        ViewContentCopier.CopyViewProperties(srcView, matchingView);
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
                Trace.WriteLine(
                    $"[DEBUG] Drafting '{srcView.Name}': starting doc-to-doc copy.");

                // Snapshot existing view IDs to find the new one
                var existingIds = new HashSet<int>(
                    new FilteredElementCollector(target)
                        .OfClass(typeof(View))
                        .Select(v => v.Id.IntegerValue));

                Trace.WriteLine(
                    $"[DEBUG] Drafting '{srcView.Name}': "
                    + $"target has {existingIds.Count} existing views before copy.");

                // Check source view contents BEFORE copying
                var srcContents = new FilteredElementCollector(source, srcView.Id)
                    .WhereElementIsNotElementType()
                    .ToElementIds();
                Trace.WriteLine(
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

                        Trace.WriteLine(
                            $"[DEBUG] Drafting '{srcView.Name}': "
                            + $"CopyElements returned {copied?.Count ?? 0} ids.");

                        if (copied != null && copied.Count > 0)
                        {
                            foreach (var cid in copied)
                            {
                                Element el = target.GetElement(cid);
                                bool isNew = !existingIds.Contains(cid.IntegerValue);
                                Trace.WriteLine(
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
                                Trace.WriteLine(
                                    $"[DEBUG] Drafting '{srcView.Name}': "
                                    + $"new view has {newContents.Count} elements before commit.");

                                t.Commit();
                                Trace.WriteLine(
                                    $"[DEBUG] Drafting '{srcView.Name}': committed OK.");
                            }
                            else
                            {
                                t.RollBack();
                                Trace.WriteLine(
                                    $"[DEBUG] Drafting '{srcView.Name}': "
                                    + "no new View found in copied ids — rolled back.");
                            }
                        }
                        else
                        {
                            t.RollBack();
                            Trace.WriteLine(
                                $"[DEBUG] Drafting '{srcView.Name}': "
                                + "CopyElements returned null or empty — rolled back.");
                        }
                    }
                    catch (Exception ex)
                    {
                        t.RollBack();
                        Trace.WriteLine(
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
                    Trace.WriteLine(
                        $"[DEBUG] Drafting '{srcView.Name}': "
                        + $"final view '{newView.Name}' has {finalContents.Count} elements after commit.");
                }
                else
                {
                    Trace.WriteLine(
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
                                ViewContentCopier.CopyViewProperties(srcView, newView);
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
                                Trace.WriteLine(
                                    $"[DBG-PLAN] '{srcView.Name}': "
                                    + $"doc-to-doc copy OK, id={newId.IntegerValue}");
                            }
                            else
                            {
                                t.RollBack();
                                Trace.WriteLine(
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
                            Trace.WriteLine(
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

                        Trace.WriteLine(
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
                        var contentIds = ViewContentCopier.GetViewContents(
                            source, srcView, opts);

                        Trace.WriteLine(
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
                            Trace.WriteLine(
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

                            Trace.WriteLine(
                                $"[DBG-CONTENT] '{srcView.Name}': "
                                + $"after filter (view-owned only): {contentIds.Count} elements.");

                            var copiedIds = ElementTransformUtils.CopyElements(
                                srcView,
                                contentIds,
                                newView,
                                null,
                                cpOpts);

                            Trace.WriteLine(
                                $"[DBG-CONTENT] '{srcView.Name}': "
                                + $"CopyElements returned {copiedIds?.Count ?? 0} elements.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Trace.WriteLine(
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
                Trace.WriteLine(
                    $"[DBG-COSV] '{srcView.Name}': skipped content copy "
                    + $"(handled by ForceCreate, type={srcView.ViewType})");
            }

            return newView.Id;
        }


        // ══════════════════════════════════════════════════════
        //  FORCE-CREATE HELPERS
        // ══════════════════════════════════════════════════════

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
                        Trace.WriteLine(
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
                    ViewContentCopier.CopyViewProperties(srcPlan, newView);

                    t.Commit();

                    Trace.WriteLine(
                        $"[DBG-FP] '{srcPlan.Name}': created plan OK "
                        + $"id={newView.Id.IntegerValue} on level '{tgtLevel.Name}'");
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    Trace.WriteLine(
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
                    Trace.WriteLine(
                        $"[DBG-SEC1] '{srcSection.Name}': "
                        + $"src.Origin=({srcSection.Origin.X:F4},{srcSection.Origin.Y:F4},{srcSection.Origin.Z:F4}) "
                        + $"src.ViewDir=({srcSection.ViewDirection.X:F4},{srcSection.ViewDirection.Y:F4},{srcSection.ViewDirection.Z:F4}) "
                        + $"src.UpDir=({srcSection.UpDirection.X:F4},{srcSection.UpDirection.Y:F4},{srcSection.UpDirection.Z:F4}) "
                        + $"src.CropOrigin=({srcSection.CropBox.Transform.Origin.X:F4},"
                        + $"{srcSection.CropBox.Transform.Origin.Y:F4},"
                        + $"{srcSection.CropBox.Transform.Origin.Z:F4})");

                    // ── DEBUG 2: donor view BEFORE overwrite ──
                    Trace.WriteLine(
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
                    Trace.WriteLine(
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
            Trace.WriteLine(
                $"[DBG-SEC4] '{srcSection.Name}': "
                + $"postCommit.Origin=({newView.Origin.X:F4},{newView.Origin.Y:F4},{newView.Origin.Z:F4}) "
                + $"postCommit.ViewDir=({newView.ViewDirection.X:F4},{newView.ViewDirection.Y:F4},{newView.ViewDirection.Z:F4}) "
                + $"postCommit.CropOrigin=({newView.CropBox.Transform.Origin.X:F4},"
                + $"{newView.CropBox.Transform.Origin.Y:F4},"
                + $"{newView.CropBox.Transform.Origin.Z:F4})");
            XYZ srcOrigin  = srcSection.CropBox.Transform.Origin;
            XYZ destOrigin = newView.CropBox.Transform.Origin;
            XYZ delta      = srcOrigin - destOrigin;

            Trace.WriteLine(
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
                    Trace.WriteLine(
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

                            double dot   = srcDirXY.DotProduct(newDirXY);
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

                                Trace.WriteLine(
                                    $"[DBG-ROT] '{srcSection.Name}': "
                                    + $"rotated {angle * 180.0 / Math.PI:F2}° "
                                    + $"srcDir=({srcDir.X:F4},{srcDir.Y:F4},{srcDir.Z:F4}) "
                                    + $"donorDir=({newDir.X:F4},{newDir.Y:F4},{newDir.Z:F4})");
                            }

                            t.Commit();

                            // ── DEBUG: final position check ──
                            Trace.WriteLine(
                                $"[DBG-FINAL] '{srcSection.Name}': "
                                + $"final.Origin=({newView.Origin.X:F4},{newView.Origin.Y:F4},{newView.Origin.Z:F4}) "
                                + $"final.ViewDir=({newView.ViewDirection.X:F4},{newView.ViewDirection.Y:F4},{newView.ViewDirection.Z:F4}) "
                                + $"expected.Origin=({srcSection.Origin.X:F4},{srcSection.Origin.Y:F4},{srcSection.Origin.Z:F4})");
                        }
                        catch (Exception ex)
                        {
                            t.RollBack();
                            Trace.WriteLine(
                                $"[DBG-SEC] '{srcSection.Name}': "
                                + $"move/rotate failed: {ex.Message}");
                        }
                    }
                }
                else
                {
                    Trace.WriteLine(
                        $"[DEBUG-SEC] '{srcSection.Name}': "
                        + "new OST_Viewers element not found.");
                }
            }

            // ── STEP 5: Copy scale/properties (NO CropBox here) ──
            // CopyViewProperties sets CropBox again — we skip that
            // by calling only the individual properties we need.
            try
            {
                newView.Scale = srcSection.Scale;
            }
            catch
            {
                try
                {
                    var srcParam = srcSection.get_Parameter(BuiltInParameter.VIEW_SCALE_PULLDOWN_METRIC)
                                ?? srcSection.get_Parameter(BuiltInParameter.VIEW_SCALE);
                    var dstParam = newView.get_Parameter(BuiltInParameter.VIEW_SCALE_PULLDOWN_METRIC)
                                ?? newView.get_Parameter(BuiltInParameter.VIEW_SCALE);
                    if (srcParam != null && dstParam != null && !dstParam.IsReadOnly)
                        dstParam.Set(srcParam.AsInteger());
                }
                catch { }
            }
            try { newView.DetailLevel  = srcSection.DetailLevel;  } catch { }
            try { newView.DisplayStyle = srcSection.DisplayStyle; } catch { }
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
                    var contentIds = ViewContentCopier.GetViewContents(
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

                        Trace.WriteLine(
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


        // ══════════════════════════════════════════════════════
        //  SMALL UTILITIES
        // ══════════════════════════════════════════════════════

        private static Level FindClosestLevel(Document doc, double elevation)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => Math.Abs(l.Elevation - elevation))
                .FirstOrDefault();
        }

        internal static string GetUniqueViewName(Document doc, string desired)
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
    }
}
