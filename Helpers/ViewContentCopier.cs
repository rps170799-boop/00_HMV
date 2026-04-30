using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace HMVTools
{
    public static class ViewContentCopier
    {
        // View types that are entirely 2D by definition — no ViewSpecific filter needed.
        private static bool IsFlat2DView(View view)
        {
            return view.ViewType == ViewType.DraftingView
                || view.ViewType == ViewType.Legend
                || view.ViewType == ViewType.DrawingSheet;
        }

        // View types where model geometry is visible — ViewSpecific guard applies.
        private static bool IsSpatialView(View view)
        {
            return view.ViewType == ViewType.FloorPlan
                || view.ViewType == ViewType.CeilingPlan
                || view.ViewType == ViewType.Section
                || view.ViewType == ViewType.Detail
                || view.ViewType == ViewType.Elevation
                || view.ViewType == ViewType.ThreeD;
        }

        /// <summary>
        /// Collects all element IDs owned by a view, filtering out
        /// viewports, extent elements, guide grids, and optionally
        /// title blocks and schedules based on options.
        /// For spatial views (plan/section/elevation/3D), only view-owned
        /// annotations are collected unless IncludeModelElements is true.
        /// Mirrors pyRevit get_view_contents.
        /// </summary>
        public static List<ElementId> GetViewContents(
            Document doc, View view, SheetCopyOptions options)
        {
            var result = new List<ElementId>();

            // Whether to enforce the ViewSpecific guard
            bool enforceViewSpecific = !options.IncludeModelElements
                                       && IsSpatialView(view);

            var elements = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .ToElements();

            foreach (Element el in elements)
            {
                if (el == null) continue;

                // For spatial views: skip elements that are merely visible
                // (not owned by this view). Prevents dragging model geometry
                // along with its hosted dimensions.
                if (enforceViewSpecific && !el.ViewSpecific)
                    continue;

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
                        System.StringComparison.OrdinalIgnoreCase) >= 0)
                    continue;

                // Skip view references
                if (el.Category != null
                    && string.Equals(el.Category.Name, "views",
                        System.StringComparison.OrdinalIgnoreCase))
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
            // Scale (with fallback for custom/non-standard values)
            try
            {
                destView.Scale = sourceView.Scale;
            }
            catch
            {
                try
                {
                    var srcParam = sourceView.get_Parameter(BuiltInParameter.VIEW_SCALE_PULLDOWN_METRIC)
                                ?? sourceView.get_Parameter(BuiltInParameter.VIEW_SCALE);
                    var dstParam = destView.get_Parameter(BuiltInParameter.VIEW_SCALE_PULLDOWN_METRIC)
                                ?? destView.get_Parameter(BuiltInParameter.VIEW_SCALE);
                    if (srcParam != null && dstParam != null && !dstParam.IsReadOnly)
                        dstParam.Set(srcParam.AsInteger());
                }
                catch { }
            }

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
    }
}
