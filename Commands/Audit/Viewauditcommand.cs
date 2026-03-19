using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class ViewAuditCommand : IExternalCommand
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
                // ── 1. Build View → Sheets map via Viewports ───
                var viewports = new FilteredElementCollector(doc)
                    .OfClass(typeof(Viewport))
                    .Cast<Viewport>()
                    .ToList();

                // viewId → list of "SheetNumber - SheetName"
                var viewToSheets =
                    new Dictionary<ElementId, List<string>>();

                foreach (var vp in viewports)
                {
                    ElementId viewId = vp.ViewId;
                    ViewSheet sheet =
                        doc.GetElement(vp.SheetId) as ViewSheet;
                    if (sheet == null) continue;

                    string info =
                        $"{sheet.SheetNumber} - {sheet.Name}";

                    if (!viewToSheets.ContainsKey(viewId))
                        viewToSheets[viewId] = new List<string>();

                    if (!viewToSheets[viewId].Contains(info))
                        viewToSheets[viewId].Add(info);
                }

                if (viewToSheets.Count == 0)
                {
                    TaskDialog.Show("HMV Tools",
                        "No views placed on sheets were found.");
                    return Result.Succeeded;
                }

                // ── 2. Build entries for the window ────────────
                var entries = new List<ViewAuditEntry>();

                foreach (var kvp in viewToSheets)
                {
                    View view = doc.GetElement(kvp.Key) as View;
                    if (view == null) continue;

                    string viewType = GetFriendlyViewType(view);

                    entries.Add(new ViewAuditEntry
                    {
                        ElementId = view.Id.IntegerValue,
                        ViewType = viewType,
                        OriginalName = view.Name,
                        NewName = view.Name,
                        Sheets = string.Join(";  ", kvp.Value),
                        SheetCount = kvp.Value.Count
                    });
                }

                // Sort: by Type, then by Name
                entries = entries
                    .OrderBy(e => e.ViewType)
                    .ThenBy(e => e.OriginalName)
                    .ToList();

                // ── 3. Collect ALL view names for conflict check
                var allViewNames = new HashSet<string>(
                    new FilteredElementCollector(doc)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .Where(v => !v.IsTemplate)
                        .Select(v => v.Name));

                // ── 4. Show window ─────────────────────────────
                var window = new ViewAuditWindow(
                    entries, allViewNames);

                bool? ok = window.ShowDialog();
                if (ok != true || window.Results == null)
                    return Result.Cancelled;

                // ── 5. Apply renames ───────────────────────────
                var toRename = window.Results
                    .Where(e => e.NameChanged && !e.HasConflict)
                    .ToList();

                if (toRename.Count == 0)
                {
                    TaskDialog.Show("HMV Tools",
                        "No views were renamed.");
                    return Result.Succeeded;
                }

                int renamed = 0;
                int skipped = 0;
                var errors = new List<string>();

                using (Transaction tx = new Transaction(
                    doc, "HMV – Rename Views"))
                {
                    tx.Start();

                    foreach (var entry in toRename)
                    {
                        try
                        {
                            ElementId id =
                                new ElementId(entry.ElementId);
                            View view = doc.GetElement(id) as View;
                            if (view == null)
                            {
                                skipped++;
                                continue;
                            }

                            view.Name = entry.NewName;
                            renamed++;
                        }
                        catch (Exception ex)
                        {
                            skipped++;
                            errors.Add(
                                $"  • {entry.OriginalName} → "
                                + $"{entry.NewName}: {ex.Message}");
                        }
                    }

                    tx.Commit();
                }

                // ── 6. Summary ─────────────────────────────────
                int conflicts = window.Results
                    .Count(e => e.HasConflict);

                string summary =
                    $"Renamed: {renamed} view(s)\n"
                    + $"Skipped (conflict): {conflicts}\n"
                    + $"Errors: {skipped}";

                if (errors.Count > 0)
                    summary += "\n\nErrors:\n"
                        + string.Join("\n", errors);

                TaskDialog.Show("HMV Tools – View Audit", summary);

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions
                .OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("HMV Tools – Error", ex.ToString());
                return Result.Failed;
            }
        }

        /// <summary>
        /// Returns a readable view type string.
        /// </summary>
        private string GetFriendlyViewType(View view)
        {
            switch (view.ViewType)
            {
                case ViewType.FloorPlan: return "Floor Plan";
                case ViewType.CeilingPlan: return "Ceiling Plan";
                case ViewType.EngineeringPlan: return "Structural Plan";
                case ViewType.AreaPlan: return "Area Plan";
                case ViewType.Section: return "Section";
                case ViewType.Elevation: return "Elevation";
                case ViewType.Detail: return "Detail";
                case ViewType.ThreeD: return "3D View";
                case ViewType.Legend: return "Legend";
                case ViewType.DraftingView: return "Drafting View";
                case ViewType.Schedule: return "Schedule";
                case ViewType.Rendering: return "Rendering";
                case ViewType.Walkthrough: return "Walkthrough";
                default: return view.ViewType.ToString();
            }
        }
    }
}