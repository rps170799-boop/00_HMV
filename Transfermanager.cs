using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HMVTools
{
    // ── Duplicate type handler ─────────────────────────────────

    /// <summary>
    /// Forces destination types when names collide, preventing
    /// the ".1" suffix pollution on company-standard families.
    /// </summary>
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
        public int AnnotationsCopied { get; set; }
        public int RefMarkersCreated { get; set; }
        public List<string> Warnings { get; set; } = new List<string>();
        public List<string> Errors { get; set; } = new List<string>();

        public string BuildReport()
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("═══  MIGRATION REPORT  ═══");
            sb.AppendLine();
            sb.AppendLine($"  Elements copied:       {ElementsCopied}");
            sb.AppendLine($"  Views created:         {ViewsCreated}");
            sb.AppendLine($"  Annotations migrated:  {AnnotationsCopied}");
            sb.AppendLine($"  Ref markers created:   {RefMarkersCreated}");

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

    /// <summary>
    /// Encapsulates all logic for inter-document element and view
    /// migration using Shared Coordinates alignment.
    /// </summary>
    public static class TransferManager
    {
        // ═══════════════════════════════════════════════════════
        //  1. SHARED COORDINATES TRANSFORM
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Computes the transform that maps source-document world
        /// coordinates to target-document world coordinates via
        /// shared coordinates:  T_final = T_target⁻¹ · T_source.
        /// </summary>
        /// <returns>
        /// The transform to apply when copying elements, or
        /// <c>Transform.Identity</c> if both share the same origin.
        /// </returns>
        public static Transform ComputeSharedCoordinateTransform(
            Document source, Document target)
        {
            Transform tSrc = source.ActiveProjectLocation
                                   .GetTotalTransform();
            Transform tTgt = target.ActiveProjectLocation
                                   .GetTotalTransform();

            return tTgt.Inverse.Multiply(tSrc);
        }

        /// <summary>
        /// Validates that the transform is reasonable (no NaN,
        /// no extreme translations).  Returns false with a
        /// message if it looks like shared coordinates are not
        /// set up.
        /// </summary>
        public static bool ValidateTransform(
            Transform t, out string warning)
        {
            warning = null;

            // Check for NaN in origin
            if (double.IsNaN(t.Origin.X) ||
                double.IsNaN(t.Origin.Y) ||
                double.IsNaN(t.Origin.Z))
            {
                warning = "Transform contains NaN — shared coordinates "
                        + "may not be configured.";
                return false;
            }

            // Warn if translation exceeds 10 km (≈32808 ft)
            double dist = t.Origin.GetLength();
            if (dist > 32808)
            {
                warning = $"Translation is {dist:F0} ft (~{dist * 0.3048:F0} m). "
                        + "Shared coordinates may be misconfigured.";
                // Allow to proceed, but warn
            }

            return true;
        }

        // ═══════════════════════════════════════════════════════
        //  2. COPY 3D ELEMENTS
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Copies model elements from source to target using the
        /// shared-coordinate transform.  Excludes view-specific
        /// elements and OST_Viewers.
        /// </summary>
        public static ICollection<ElementId> CopyModelElements(
            Document source,
            Document target,
            ICollection<ElementId> elementIds,
            Transform coordTransform,
            TransferResult result)
        {
            // Filter out view-specific and viewer elements
            var validIds = new List<ElementId>();
            foreach (var id in elementIds)
            {
                Element el = source.GetElement(id);
                if (el == null) continue;

                // Skip view-specific elements (annotations, etc.)
                if (el.OwnerViewId != ElementId.InvalidElementId)
                {
                    result.Warnings.Add(
                        $"Skipped view-specific element: {el.Name} (Id {id.IntegerValue})");
                    continue;
                }

                // Skip viewers (section markers, etc.)
                if (el.Category != null &&
                    (BuiltInCategory)el.Category.Id.IntegerValue
                        == BuiltInCategory.OST_Viewers)
                {
                    continue;
                }

                validIds.Add(id);
            }

            if (validIds.Count == 0)
                return new List<ElementId>();

            var opts = new CopyPasteOptions();
            opts.SetDuplicateTypeNamesHandler(
                new UseDestinationTypesHandler());

            ICollection<ElementId> copied;
            using (var t = new Transaction(target, "Copy Model Elements"))
            {
                t.Start();
                try
                {
                    copied = ElementTransformUtils.CopyElements(
                        source, validIds, target,
                        coordTransform, opts);
                    t.Commit();
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    result.Errors.Add($"CopyElements failed: {ex.Message}");
                    return new List<ElementId>();
                }
            }

            result.ElementsCopied = copied.Count;
            return copied;
        }

        // ═══════════════════════════════════════════════════════
        //  3. VIEW RECONSTRUCTION
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Dispatches view recreation based on view type.
        /// Returns the new view in the target document, or null.
        /// </summary>
        public static View RecreateView(
            Document source,
            Document target,
            View sourceView,
            Transform coordTransform,
            TransferResult result)
        {
            try
            {
                if (sourceView is ViewPlan vp)
                    return RecreateViewPlan(source, target, vp, coordTransform, result);

                if (sourceView is ViewSection vs)
                    return RecreateViewSection(source, target, vs, coordTransform, result);

                if (sourceView.ViewType == ViewType.DraftingView)
                    return TransferDraftingView(source, target, sourceView, result);

                if (sourceView.ViewType == ViewType.Legend)
                    return TransferLegendView(source, target, sourceView, result);

                result.Warnings.Add(
                    $"Unsupported view type '{sourceView.ViewType}' for '{sourceView.Name}'.");
                return null;
            }
            catch (Exception ex)
            {
                result.Errors.Add(
                    $"Failed to recreate view '{sourceView.Name}': {ex.Message}");
                return null;
            }
        }

        // ── 3a. ViewPlan ───────────────────────────────────────

        private static ViewPlan RecreateViewPlan(
            Document source, Document target,
            ViewPlan sourceView, Transform coordTransform,
            TransferResult result)
        {
            // Find a matching ViewFamilyType in target
            ElementId vftId = FindViewFamilyType(target, sourceView.ViewType);
            if (vftId == null || vftId == ElementId.InvalidElementId)
            {
                result.Errors.Add(
                    $"No ViewFamilyType for '{sourceView.ViewType}' in target.");
                return null;
            }

            // Get the source level elevation and find closest in target
            Level srcLevel = source.GetElement(sourceView.GenLevel.Id) as Level;
            double srcElevation = srcLevel?.Elevation ?? 0;

            // Transform the elevation to target coordinate space
            XYZ srcPt = new XYZ(0, 0, srcElevation);
            XYZ tgtPt = coordTransform.OfPoint(srcPt);
            double targetElevation = tgtPt.Z;

            Level tgtLevel = FindClosestLevel(target, targetElevation);
            if (tgtLevel == null)
            {
                result.Errors.Add(
                    $"No levels found in target document for view '{sourceView.Name}'.");
                return null;
            }

            ViewPlan newView;
            using (var t = new Transaction(target, "Create ViewPlan"))
            {
                t.Start();
                newView = ViewPlan.Create(target, vftId, tgtLevel.Id);

                // Name — handle duplicates
                newView.Name = GetUniqueViewName(target, sourceView.Name);

                // Crop box — transform from source to target coords
                if (sourceView.CropBoxActive)
                {
                    BoundingBoxXYZ srcCrop = sourceView.CropBox;
                    BoundingBoxXYZ tgtCrop = TransformCropBox(
                        srcCrop, coordTransform);
                    newView.CropBox = tgtCrop;
                    newView.CropBoxActive = true;
                    newView.CropBoxVisible = sourceView.CropBoxVisible;
                }

                // View range — copy the PlanViewRange offsets
                try
                {
                    PlanViewRange srcRange = sourceView.GetViewRange();
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
                catch
                {
                    result.Warnings.Add(
                        $"Could not copy view range for '{sourceView.Name}'.");
                }

                // Match view template by name if available
                ApplyViewTemplateByName(source, target, sourceView, newView, result);

                t.Commit();
            }

            target.Regenerate();
            result.ViewsCreated++;
            return newView;
        }

        // ── 3b. ViewSection ────────────────────────────────────

        private static ViewSection RecreateViewSection(
            Document source, Document target,
            ViewSection sourceView, Transform coordTransform,
            TransferResult result)
        {
            ElementId vftId = FindViewFamilyType(target, sourceView.ViewType);
            if (vftId == null || vftId == ElementId.InvalidElementId)
            {
                result.Errors.Add(
                    $"No ViewFamilyType for '{sourceView.ViewType}' in target.");
                return null;
            }

            // Build the section box in target coordinates
            BoundingBoxXYZ srcBox = sourceView.CropBox;
            BoundingBoxXYZ sectionBox = TransformSectionBox(
                srcBox, coordTransform);

            ViewSection newView;
            using (var t = new Transaction(target, "Create ViewSection"))
            {
                t.Start();
                newView = ViewSection.CreateSection(
                    target, vftId, sectionBox);

                newView.Name = GetUniqueViewName(target, sourceView.Name);

                newView.CropBoxActive = sourceView.CropBoxActive;
                newView.CropBoxVisible = sourceView.CropBoxVisible;

                // Far clip
                try
                {
                    Parameter farClip = newView.get_Parameter(
                        BuiltInParameter.VIEWER_BOUND_OFFSET_FAR);
                    Parameter srcFarClip = sourceView.get_Parameter(
                        BuiltInParameter.VIEWER_BOUND_OFFSET_FAR);
                    if (farClip != null && srcFarClip != null
                        && !farClip.IsReadOnly)
                    {
                        farClip.Set(srcFarClip.AsDouble());
                    }
                }
                catch { /* non-critical */ }

                ApplyViewTemplateByName(source, target, sourceView, newView, result);
                t.Commit();
            }

            target.Regenerate();
            result.ViewsCreated++;
            return newView;
        }

        // ── 3c. Drafting Views ─────────────────────────────────

        private static View TransferDraftingView(
            Document source, Document target,
            View sourceView, TransferResult result)
        {
            ElementId vftId = FindViewFamilyType(target, ViewType.DraftingView);
            if (vftId == null || vftId == ElementId.InvalidElementId)
            {
                result.Errors.Add("No DraftingView ViewFamilyType in target.");
                return null;
            }

            ViewDrafting newView;
            using (var t = new Transaction(target, "Create Drafting View"))
            {
                t.Start();
                newView = ViewDrafting.Create(target, vftId);
                newView.Name = GetUniqueViewName(target, sourceView.Name);
                newView.Scale = sourceView.Scale;
                t.Commit();
            }

            target.Regenerate();

            // Copy all owned elements (2D content)
            CopyOwnedElements(source, target, sourceView, newView,
                Transform.Identity, result);

            result.ViewsCreated++;
            return newView;
        }

        // ── 3d. Legend Views ───────────────────────────────────

        private static View TransferLegendView(
            Document source, Document target,
            View sourceView, TransferResult result)
        {
            ElementId vftId = FindViewFamilyType(target, ViewType.Legend);
            if (vftId == null || vftId == ElementId.InvalidElementId)
            {
                result.Errors.Add("No Legend ViewFamilyType in target.");
                return null;
            }

            View newView;
            using (var t = new Transaction(target, "Create Legend View"))
            {
                t.Start();
                newView = ViewDrafting.Create(target, vftId);
                newView.Name = GetUniqueViewName(target, sourceView.Name);
                newView.Scale = sourceView.Scale;
                t.Commit();
            }

            target.Regenerate();

            // Pre-check: warn about missing family types for
            // LegendComponent elements
            var legendComps = new FilteredElementCollector(source, sourceView.Id)
                .OfCategory(BuiltInCategory.OST_LegendComponents)
                .ToElementIds();

            if (legendComps.Count > 0)
            {
                var missingTypes = CheckLegendFamilyTypes(
                    source, target, legendComps);
                foreach (var mt in missingTypes)
                    result.Warnings.Add(
                        $"Legend type missing in target: {mt}");
            }

            CopyOwnedElements(source, target, sourceView, newView,
                Transform.Identity, result);

            result.ViewsCreated++;
            return newView;
        }

        // ═══════════════════════════════════════════════════════
        //  4. ANNOTATION MIGRATION (2-STEP FOR REVIT 2023)
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Copies view-specific elements (detail lines, text,
        /// dimensions, filled regions) from sourceView into
        /// targetView using the view-to-view CopyElements overload.
        /// Revit 2023 requires the target view to already exist
        /// and the document to be regenerated before this call.
        /// </summary>
        public static void CopyViewAnnotations(
            Document sourceDoc, Document targetDoc,
            View sourceView, View targetView,
            Transform viewTransform,
            TransferResult result)
        {
            // Collect all owned elements, excluding viewers
            var collector = new FilteredElementCollector(
                sourceDoc, sourceView.Id);

            var ownedIds = collector
                .WhereElementIsNotElementType()
                .Where(e =>
                {
                    if (e.Category == null) return true;
                    var bic = (BuiltInCategory)e.Category.Id.IntegerValue;
                    // Exclude section/callout/elevation markers
                    return bic != BuiltInCategory.OST_Viewers
                        && bic != BuiltInCategory.OST_SectionBox;
                })
                .Select(e => e.Id)
                .ToList();

            if (ownedIds.Count == 0) return;

            var opts = new CopyPasteOptions();
            opts.SetDuplicateTypeNamesHandler(
                new UseDestinationTypesHandler());

            using (var t = new Transaction(targetDoc, "Copy Annotations"))
            {
                t.Start();
                try
                {
                    var copied = ElementTransformUtils.CopyElements(
                        sourceView, ownedIds,
                        targetView, viewTransform, opts);

                    result.AnnotationsCopied += copied.Count;
                }
                catch (Exception ex)
                {
                    result.Warnings.Add(
                        $"Partial annotation copy for '{sourceView.Name}': "
                        + ex.Message);
                }
                t.Commit();
            }
        }

        // ═══════════════════════════════════════════════════════
        //  5. REFERENCE MARKER RECREATION
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Identifies section/callout/elevation markers in the
        /// source view and notes them for recreation.  Full
        /// recreation of markers requires placing new ViewSection
        /// references which is handled per-view.
        /// </summary>
        public static List<string> IdentifyReferenceMarkers(
            Document source, View sourceView)
        {
            var markers = new List<string>();

            try
            {
                var refSections = sourceView.GetReferenceSections();
                if (refSections != null)
                    foreach (var id in refSections)
                    {
                        var el = source.GetElement(id);
                        if (el != null)
                            markers.Add($"Section ref: {el.Name} (Id {id.IntegerValue})");
                    }
            }
            catch { /* Method may not be available on all view types */ }

            try
            {
                var refCallouts = sourceView.GetReferenceCallouts();
                if (refCallouts != null)
                    foreach (var id in refCallouts)
                    {
                        var el = source.GetElement(id);
                        if (el != null)
                            markers.Add($"Callout ref: {el.Name} (Id {id.IntegerValue})");
                    }
            }
            catch { }

            try
            {
                var refElevs = sourceView.GetReferenceElevations();
                if (refElevs != null)
                    foreach (var id in refElevs)
                    {
                        var el = source.GetElement(id);
                        if (el != null)
                            markers.Add($"Elevation ref: {el.Name} (Id {id.IntegerValue})");
                    }
            }
            catch { }

            return markers;
        }

        // ═══════════════════════════════════════════════════════
        //  HELPERS
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Finds the closest Level in the document by absolute
        /// elevation difference (in feet).
        /// </summary>
        public static Level FindClosestLevel(Document doc, double elevationFeet)
        {
            var levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .ToList();

            if (levels.Count == 0) return null;

            return levels
                .OrderBy(l => Math.Abs(l.Elevation - elevationFeet))
                .First();
        }

        /// <summary>
        /// Finds a ViewFamilyType in the target document matching
        /// the given ViewType.
        /// </summary>
        private static ElementId FindViewFamilyType(
            Document doc, ViewType viewType)
        {
            // Map ViewType to ViewFamily
            ViewFamily vf;
            switch (viewType)
            {
                case ViewType.FloorPlan: vf = ViewFamily.FloorPlan; break;
                case ViewType.CeilingPlan: vf = ViewFamily.CeilingPlan; break;
                case ViewType.Section: vf = ViewFamily.Section; break;
                case ViewType.Detail: vf = ViewFamily.Detail; break;
                case ViewType.DraftingView: vf = ViewFamily.Drafting; break;
                case ViewType.Legend: vf = ViewFamily.Legend; break;
                case ViewType.Elevation: vf = ViewFamily.Elevation; break;
                default:
                    return ElementId.InvalidElementId;
            }

            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .Where(t => t.ViewFamily == vf)
                .Select(t => t.Id)
                .FirstOrDefault() ?? ElementId.InvalidElementId;
        }

        /// <summary>
        /// Transforms a crop box (BoundingBoxXYZ) for a plan view
        /// by mapping its Min/Max corners through the shared
        /// coordinate transform.
        /// </summary>
        private static BoundingBoxXYZ TransformCropBox(
            BoundingBoxXYZ srcCrop, Transform coordTransform)
        {
            // The crop box transform defines the view's local CS
            Transform srcT = srcCrop.Transform;

            // Transform the 8 corners and find new AABB in target
            // view coords — but for plans we can do a simpler
            // approach: transform origin and keep local extents
            // since plan orientation doesn't usually change.

            var newCrop = new BoundingBoxXYZ();

            // Transform the view's local coordinate system
            XYZ newOrigin = coordTransform.OfPoint(srcT.Origin);
            XYZ newBasisX = coordTransform.OfVector(srcT.BasisX);
            XYZ newBasisY = coordTransform.OfVector(srcT.BasisY);
            XYZ newBasisZ = coordTransform.OfVector(srcT.BasisZ);

            Transform newT = Transform.Identity;
            newT.Origin = newOrigin;
            newT.BasisX = newBasisX.Normalize();
            newT.BasisY = newBasisY.Normalize();
            newT.BasisZ = newBasisZ.Normalize();

            newCrop.Transform = newT;
            newCrop.Min = srcCrop.Min; // local extents unchanged
            newCrop.Max = srcCrop.Max;
            newCrop.Enabled = srcCrop.Enabled;

            return newCrop;
        }

        /// <summary>
        /// Transforms a section box for ViewSection.CreateSection.
        /// The section box Transform defines the section's
        /// orientation; Min/Max define the depth and extents.
        /// </summary>
        private static BoundingBoxXYZ TransformSectionBox(
            BoundingBoxXYZ srcBox, Transform coordTransform)
        {
            Transform srcT = srcBox.Transform;

            var newBox = new BoundingBoxXYZ();

            XYZ newOrigin = coordTransform.OfPoint(srcT.Origin);
            XYZ newBasisX = coordTransform.OfVector(srcT.BasisX).Normalize();
            XYZ newBasisY = coordTransform.OfVector(srcT.BasisY).Normalize();
            XYZ newBasisZ = coordTransform.OfVector(srcT.BasisZ).Normalize();

            Transform newT = Transform.Identity;
            newT.Origin = newOrigin;
            newT.BasisX = newBasisX;
            newT.BasisY = newBasisY;
            newT.BasisZ = newBasisZ;

            newBox.Transform = newT;
            newBox.Min = srcBox.Min;
            newBox.Max = srcBox.Max;
            newBox.Enabled = true;

            return newBox;
        }

        /// <summary>
        /// Attempts to find a view template in the target document
        /// with the same name as the source view's template, and
        /// applies it.
        /// </summary>
        private static void ApplyViewTemplateByName(
            Document source, Document target,
            View sourceView, View targetView,
            TransferResult result)
        {
            if (sourceView.ViewTemplateId == ElementId.InvalidElementId)
                return;

            View srcTemplate = source.GetElement(
                sourceView.ViewTemplateId) as View;
            if (srcTemplate == null) return;

            string templateName = srcTemplate.Name;

            var tgtTemplate = new FilteredElementCollector(target)
                .OfClass(typeof(View))
                .Cast<View>()
                .FirstOrDefault(v => v.IsTemplate
                    && v.Name == templateName);

            if (tgtTemplate != null)
            {
                targetView.ViewTemplateId = tgtTemplate.Id;
            }
            else
            {
                result.Warnings.Add(
                    $"View template '{templateName}' not found in target. "
                    + $"View '{targetView.Name}' has no template.");
            }
        }

        /// <summary>
        /// Generates a unique view name in the target document,
        /// appending " (migrated)" or " (migrated 2)" if needed.
        /// </summary>
        private static string GetUniqueViewName(
            Document doc, string desiredName)
        {
            var existingNames = new HashSet<string>(
                new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => !v.IsTemplate)
                    .Select(v => v.Name));

            if (!existingNames.Contains(desiredName))
                return desiredName;

            string baseName = desiredName + " (migrated)";
            if (!existingNames.Contains(baseName))
                return baseName;

            int counter = 2;
            while (existingNames.Contains($"{baseName} {counter}"))
                counter++;

            return $"{baseName} {counter}";
        }

        /// <summary>
        /// Copies all owned elements from a source view to a
        /// target view (used for drafting/legend views).
        /// Excludes OST_Viewers.
        /// </summary>
        private static void CopyOwnedElements(
            Document sourceDoc, Document targetDoc,
            View sourceView, View targetView,
            Transform transform, TransferResult result)
        {
            var ownedIds = new FilteredElementCollector(
                    sourceDoc, sourceView.Id)
                .WhereElementIsNotElementType()
                .Where(e =>
                {
                    if (e.Category == null) return true;
                    return (BuiltInCategory)e.Category.Id.IntegerValue
                        != BuiltInCategory.OST_Viewers;
                })
                .Select(e => e.Id)
                .ToList();

            if (ownedIds.Count == 0) return;

            var opts = new CopyPasteOptions();
            opts.SetDuplicateTypeNamesHandler(
                new UseDestinationTypesHandler());

            using (var t = new Transaction(targetDoc, "Copy View Content"))
            {
                t.Start();
                try
                {
                    var copied = ElementTransformUtils.CopyElements(
                        sourceView, ownedIds,
                        targetView, transform, opts);
                    result.AnnotationsCopied += copied.Count;
                }
                catch (Exception ex)
                {
                    result.Warnings.Add(
                        $"Partial copy for '{sourceView.Name}': {ex.Message}");
                }
                t.Commit();
            }
        }

        /// <summary>
        /// Checks which FamilyTypes referenced by LegendComponent
        /// elements are missing in the target document.
        /// </summary>
        private static List<string> CheckLegendFamilyTypes(
            Document source, Document target,
            ICollection<ElementId> legendCompIds)
        {
            var missing = new List<string>();

            // Build set of type names in target
            var targetTypeNames = new HashSet<string>(
                new FilteredElementCollector(target)
                    .WhereElementIsElementType()
                    .Select(e => e.Name));

            foreach (var id in legendCompIds)
            {
                Element el = source.GetElement(id);
                if (el == null) continue;

                ElementId typeId = el.GetTypeId();
                if (typeId == ElementId.InvalidElementId) continue;

                Element type = source.GetElement(typeId);
                if (type == null) continue;

                if (!targetTypeNames.Contains(type.Name))
                    missing.Add(type.Name);
            }

            return missing.Distinct().ToList();
        }
    }
}