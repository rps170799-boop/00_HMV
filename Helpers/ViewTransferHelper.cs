using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace HMVTools
{
    public static class ViewTransferHelper
    {
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
    }
}
