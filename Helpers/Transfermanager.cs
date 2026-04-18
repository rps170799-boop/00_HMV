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
            sb.AppendLine($"  Views transferred:        {ViewsCreated}");
            sb.AppendLine($"  Templates remapped:       {TemplatesRemapped}");
            sb.AppendLine($"  Ref markers identified:   {RefMarkersNoted}");

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
        // ═══════════════════════════════════════════════════════

        public static Dictionary<ElementId, ElementId> CopyViews(
            Document source,
            Document target,
            ICollection<ElementId> viewIds,
            Transform coordTransform,
            TransferResult result)
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
                    "Plan/Section/Elevation");
            }

            if (flatIds.Count > 0)
            {
                CopyViewBatch(source, target,
                    flatIds, Transform.Identity,
                    opts, viewMap, result,
                    "Drafting/Legend");
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
     string batchLabel)
        {
            using (var t = new Transaction(target,
                $"HMV – Copy {batchLabel} Views"))
            {
                t.Start();

                // Snapshot existing views to detect "fake" copies
                var existingViewIds = new HashSet<int>(
                    new FilteredElementCollector(target)
                        .OfClass(typeof(View))
                        .Select(v => v.Id.IntegerValue));
                int batchSuccesses = 0;
                int failures = 0;

                // Copy one view at a time so the source→target pairing is
                // unambiguous (CopyElements does NOT guarantee return order
                // matches input order). Each view is wrapped in its own
                // SubTransaction so a single failure doesn't poison the
                // parent transaction or kill the rest of the batch.
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

                    // Pre-check: use ForceCreate for all spatial views
                    // when the target already has matching content.
                    // CopyElements is unreliable for Plans (returns
                    // existing view) and Sections (misplaces origin).
                    bool needsForcePath = false;

                    if (srcV is ViewPlan srcPlan && srcPlan.GenLevel != null)
                    {
                        // Plans: force if level already exists
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
                        // Sections: always force — CopyElements
                        // doesn't reliably transform the crop box
                        needsForcePath = true;
                    }

                    using (var sub = new SubTransaction(target))
                    {
                        try
                        {
                            sub.Start();

                            if (needsForcePath)
                            {
                                // Go straight to Duplicate — no CopyElements
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
                                // Normal path: CopyElements
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
                                        // Shouldn't happen but safety net
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

                // Commit whatever succeeded. We only fail the whole batch if
                // literally nothing went through.
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
        /// <summary>
        /// Copies source view templates into the target document
        /// and assigns them to the migrated views.
        /// Duplicate names get Revit's automatic rename (e.g. "Template 1").
        /// Each unique template is copied only once and reused.
        /// </summary>
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

            // Cache: source template Id → already-copied target template Id
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

                    // Copy template only once per unique source template
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

                    // Assign to migrated view
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

            // Categories to copy, in order of likelihood to succeed
            var categories = new[]
            {
                BuiltInCategory.OST_Lines,          // detail lines
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

                    // Try view-to-view first
                    bool copied = TryCopyAnnotations(
                        srcView, tgtView, ids, opts, result,
                        $"{bic} in '{srcView.Name}'");

                    // No fallback — doc-to-doc copies cause
                    // link name conflicts and orphaned elements
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
                        $"{label}: view-to-view blocked, "
                        + "tried doc-to-doc fallback.");
                    return false;
                }
            }
        }


        // ═══════════════════════════════════════════════════════
        //  6. COPY CATEGORY GRAPHIC OVERRIDES
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Copies per-category graphic overrides from source views
        /// to target views. Remaps FillPattern and LinePattern IDs
        /// by name lookup, copying missing patterns to target.
        /// </summary>
        public static void CopyCategoryOverrides(
            Document source,
            Document target,
            Dictionary<ElementId, ElementId> viewMap,
            TransferResult result)
        {
            if (viewMap.Count == 0) return;

            // Cache pattern remapping: source PatternId → target PatternId
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

                    // Skip drafting/legend — no model categories
                    if (srcView.ViewType == ViewType.DraftingView
                        || srcView.ViewType == ViewType.Legend)
                        continue;

                    // Iterate all categories in the document
                    foreach (Category cat in source.Settings.Categories)
                    {
                        if (cat == null) continue;
                        ElementId catId = cat.Id;

                        try
                        {
                            OverrideGraphicSettings srcOgs =
                                srcView.GetCategoryOverrides(catId);

                            if (IsOverrideEmpty(srcOgs)) continue;

                            // Remap pattern IDs
                            OverrideGraphicSettings tgtOgs =
                                RemapOverride(source, target,
                                    srcOgs, fillPatternMap,
                                    linePatternMap, copyOpts);

                            tgtView.SetCategoryOverrides(catId, tgtOgs);
                            overrideCount++;
                        }
                        catch { /* category may not be overridable */ }
                    }
                }

                t.Commit();
            }

            if (overrideCount > 0)
                result.Warnings.Add(
                    $"{overrideCount} category graphic overrides applied.");
        }

        /// <summary>
        /// Creates a new OverrideGraphicSettings with all pattern
        /// and line IDs remapped from source to target document.
        /// </summary>
        private static OverrideGraphicSettings RemapOverride(
            Document source, Document target,
            OverrideGraphicSettings src,
            Dictionary<ElementId, ElementId> fillMap,
            Dictionary<ElementId, ElementId> lineMap,
            CopyPasteOptions copyOpts)
        {
            var tgt = new OverrideGraphicSettings();

            // ── Colors (direct copy, no remapping) ─────────
            tgt.SetProjectionLineColor(src.ProjectionLineColor);
            tgt.SetCutLineColor(src.CutLineColor);
            tgt.SetSurfaceForegroundPatternColor(src.SurfaceForegroundPatternColor);
            tgt.SetSurfaceBackgroundPatternColor(src.SurfaceBackgroundPatternColor);
            tgt.SetCutForegroundPatternColor(src.CutForegroundPatternColor);
            tgt.SetCutBackgroundPatternColor(src.CutBackgroundPatternColor);

            // ── Line weights (direct copy) ─────────────────
            tgt.SetProjectionLineWeight(src.ProjectionLineWeight);
            tgt.SetCutLineWeight(src.CutLineWeight);

            // ── Simple values (direct copy) ────────────────
            tgt.SetHalftone(src.Halftone);
            tgt.SetSurfaceTransparency(src.Transparency);

            // ── Line patterns (remap by name) ──────────────
            tgt.SetProjectionLinePatternId(
                RemapLinePattern(source, target,
                    src.ProjectionLinePatternId,
                    lineMap, copyOpts));

            tgt.SetCutLinePatternId(
                RemapLinePattern(source, target,
                    src.CutLinePatternId,
                    lineMap, copyOpts));

            // ── Fill patterns (remap by name) ──────────────
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

            // Get source pattern name
            FillPatternElement srcElem =
                source.GetElement(srcPatternId) as FillPatternElement;
            if (srcElem == null)
            {
                cache[srcPatternId] = ElementId.InvalidElementId;
                return ElementId.InvalidElementId;
            }

            // Try to find by name in target
            FillPatternElement tgtElem =
                FillPatternElement.GetFillPatternElementByName(
                    target, srcElem.GetFillPattern().Target,
                    srcElem.Name);

            if (tgtElem != null)
            {
                cache[srcPatternId] = tgtElem.Id;
                return tgtElem.Id;
            }

            // Not found — copy it from source
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

            // Solid line is a built-in — no remapping needed
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

            // Try name match in target
            LinePatternElement tgtElem =
                LinePatternElement.GetLinePatternElementByName(
                    target, srcElem.Name);

            if (tgtElem != null)
            {
                cache[srcPatternId] = tgtElem.Id;
                return tgtElem.Id;
            }

            // Copy from source
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

        /// <summary>
        /// Safely collects reference marker ids from a view.
        /// The getter returns List&lt;ElementId&gt; to avoid
        /// ICollection/IList conversion issues.
        /// </summary>
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
            // Find the existing view on the matching level
            Level srcLevel = source.GetElement(srcPlan.GenLevel.Id) as Level;
            if (srcLevel == null) return ElementId.InvalidElementId;

            Level tgtLevel = FindClosestLevel(target, srcLevel.Elevation);
            if (tgtLevel == null) return ElementId.InvalidElementId;

            // Find existing ViewPlan on that level
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

            // Duplicate it (WithDetailing keeps annotations if any)
            ElementId newId = existingView.Duplicate(
                ViewDuplicateOption.WithDetailing);

            ViewPlan newView = target.GetElement(newId) as ViewPlan;
            if (newView == null) return ElementId.InvalidElementId;

            newView.Name = GetUniqueViewName(target, srcPlan.Name);

            // Match crop box from source
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

            // Match view range from source
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

            // Match scale
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

            // Transform the section box to target coordinates
            Transform coordTransform =
                ComputeSharedCoordinateTransform(source, target);

            BoundingBoxXYZ srcBox = srcSection.CropBox;
            Transform srcT = srcBox.Transform;

            // Remap origin and basis vectors
            BoundingBoxXYZ tgtBox = new BoundingBoxXYZ();

            Transform newT = Transform.Identity;
            newT.Origin = coordTransform.OfPoint(srcT.Origin);
            newT.BasisX = coordTransform.OfVector(srcT.BasisX).Normalize();
            newT.BasisY = coordTransform.OfVector(srcT.BasisY).Normalize();
            newT.BasisZ = coordTransform.OfVector(srcT.BasisZ).Normalize();

            tgtBox.Transform = newT;
            tgtBox.Min = srcBox.Min;  // local extents stay the same
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
}