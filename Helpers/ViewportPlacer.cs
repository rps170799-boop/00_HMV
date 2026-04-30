using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;

namespace HMVTools
{
    // ── ViewportPlacer ─────────────────────────────────────────
    //  Viewport placement, type matching, positioning helpers.
    //  Mirrors pyRevit copy_sheet_viewports / apply_viewport_type
    //  / get_source_vport_data / correct_vport_by_bbox.
    // ──────────────────────────────────────────────────────────

    public static class ViewportPlacer
    {
        // ══════════════════════════════════════════════════════
        //  7a. COPY SHEET VIEWPORTS
        // ══════════════════════════════════════════════════════

        /// <summary>
        /// Iterates source sheet viewports, copies/updates the
        /// referenced view, then places it on the destination sheet
        /// at the original position with matching viewport type,
        /// detail number, label offset, and bounding-box correction.
        /// Mirrors pyRevit copy_sheet_viewports.
        /// </summary>
        public static void CopySheetViewports(
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

                Trace.WriteLine(
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
                    destViewId = ViewFactory.CopyOrUpdateSingleView(
                        source, target, srcView,
                        coordTransform, result, settings);

                    if (destViewId != ElementId.InvalidElementId)
                        viewMap[srcVport.ViewId] = destViewId;
                }
                Trace.WriteLine(
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
                        Trace.WriteLine(
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


        // ══════════════════════════════════════════════════════
        //  7c. VIEWPORT TYPE MATCHING
        // ══════════════════════════════════════════════════════

        /// <summary>
        /// Copies the viewport type from source to target if needed,
        /// then applies it to the new viewport.
        /// Mirrors pyRevit apply_viewport_type + copy_viewport_types.
        /// </summary>
        public static void ApplyViewportType(
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
        public static void ApplyDetailNumber(
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


        // ══════════════════════════════════════════════════════
        //  7d. VIEWPORT POSITIONING (Label, BBox correction)
        // ══════════════════════════════════════════════════════

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
        public static void ApplyLabelProperties(
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
        public static void CorrectViewportByBBox(
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
                    Trace.WriteLine(
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
        public static string GetViewSheetNumber(
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
    }
}
