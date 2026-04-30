using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace HMVTools
{
    public static class SheetHelpers
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
                return false;
            }

            return true;
        }

        // ═══════════════════════════════════════════════════════
        //  7e. SHEET GUIDE GRIDS
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Copies and assigns guide grids from source sheet to
        /// destination sheet. Mirrors pyRevit copy_sheet_guides.
        /// </summary>
        internal static void CopySheetGuides(
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
        internal static void CopySheetRevisions(
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
    }
}
