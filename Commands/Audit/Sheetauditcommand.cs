using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class SheetAuditCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // ── 1. Collect all sheets ──────────────────────
                var sheets = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .Where(s => !s.IsPlaceholder)
                    .OrderBy(s => s.SheetNumber)
                    .ToList();

                if (sheets.Count == 0)
                {
                    TaskDialog.Show("HMV Tools",
                        "No sheets found in the project.");
                    return Result.Succeeded;
                }

                // ── 2. Count viewports per sheet ───────────────
                var viewCountMap = new Dictionary<ElementId, int>();
                var viewports = new FilteredElementCollector(doc)
                    .OfClass(typeof(Viewport))
                    .Cast<Viewport>()
                    .ToList();

                foreach (var vp in viewports)
                {
                    if (!viewCountMap.ContainsKey(vp.SheetId))
                        viewCountMap[vp.SheetId] = 0;
                    viewCountMap[vp.SheetId]++;
                }

                // ── 3. Build entries ───────────────────────────
                var entries = new List<SheetAuditEntry>();
                foreach (var sheet in sheets)
                {
                    int views = 0;
                    viewCountMap.TryGetValue(sheet.Id, out views);

                    entries.Add(new SheetAuditEntry
                    {
                        ElementId = sheet.Id.IntegerValue,
                        OriginalNumber = sheet.SheetNumber,
                        NewNumber = sheet.SheetNumber,
                        OriginalName = sheet.Name,
                        NewName = sheet.Name,
                        ViewCount = views
                    });
                }

                // ── 4. Collect all sheet numbers for conflicts ─
                var allSheetNumbers = new HashSet<string>(
                    sheets.Select(s => s.SheetNumber));

                // ── 5. Show window ─────────────────────────────
                var window = new SheetAuditWindow(
                    entries, allSheetNumbers);

                bool? ok = window.ShowDialog();
                if (ok != true || window.Results == null)
                    return Result.Cancelled;

                // ── 6. Apply changes ───────────────────────────
                var toUpdate = window.Results
                    .Where(e => e.AnyChanged && !e.HasNumberConflict)
                    .ToList();

                if (toUpdate.Count == 0)
                {
                    TaskDialog.Show("HMV Tools",
                        "No sheets were modified.");
                    return Result.Succeeded;
                }

                int updated = 0;
                int skipped = 0;
                var errors = new List<string>();

                using (Transaction tx = new Transaction(
                    doc, "HMV - Rename Sheets"))
                {
                    tx.Start();

                    foreach (var entry in toUpdate)
                    {
                        try
                        {
                            ElementId id =
                                new ElementId(entry.ElementId);
                            ViewSheet sheet =
                                doc.GetElement(id) as ViewSheet;
                            if (sheet == null)
                            {
                                skipped++;
                                continue;
                            }

                            // Update number if changed
                            if (entry.NumberChanged)
                                sheet.SheetNumber = entry.NewNumber;

                            // Update name if changed
                            if (entry.NameChanged)
                                sheet.Name = entry.NewName;

                            updated++;
                        }
                        catch (Exception ex)
                        {
                            skipped++;
                            errors.Add(
                                $"  {entry.OriginalNumber} "
                                + $"({entry.OriginalName}): "
                                + ex.Message);
                        }
                    }

                    tx.Commit();
                }

                // ── 7. Summary ─────────────────────────────────
                int conflicts = window.Results
                    .Count(e => e.HasNumberConflict);

                string summary =
                    $"Updated: {updated} sheet(s)\n"
                    + $"Skipped (# conflict): {conflicts}\n"
                    + $"Errors: {skipped}";

                if (errors.Count > 0)
                    summary += "\n\nErrors:\n"
                        + string.Join("\n", errors);

                TaskDialog.Show("HMV Tools - Sheet Audit", summary);

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions
                .OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("HMV Tools - Error", ex.ToString());
                return Result.Failed;
            }
        }
    }
}