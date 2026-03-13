using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class GridLevelExtentCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // ── Collect all valid views (no templates, no sheets) ──
            List<View> allViews = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate
                         && v.ViewType != ViewType.DrawingSheet
                         && v.ViewType != ViewType.Schedule
                         && v.ViewType != ViewType.Internal
                         && v.ViewType != ViewType.Undefined
                         && v.ViewType != ViewType.ProjectBrowser
                         && v.ViewType != ViewType.SystemBrowser
                         && v.ViewType != ViewType.Legend)
                .OrderBy(v => v.ViewType.ToString())
                .ThenBy(v => v.Name)
                .ToList();

            if (allViews.Count == 0)
            {
                TaskDialog.Show("Grid & Level Extent",
                    "No valid views found in the project.");
                return Result.Cancelled;
            }

            // ── Build display items for the window (no Revit types) ──
            List<GridLevelViewItem> viewItems = new List<GridLevelViewItem>();
            for (int i = 0; i < allViews.Count; i++)
            {
                viewItems.Add(new GridLevelViewItem
                {
                    Index = i,
                    DisplayName = $"[{allViews[i].ViewType}]  {allViews[i].Name}"
                });
            }

            // ── Show WPF window ──
            GridLevelExtentWindow win = new GridLevelExtentWindow(viewItems);
            bool? result = win.ShowDialog();

            if (result != true) return Result.Cancelled;

            // ── Read user selections ──
            bool to2D = win.ConvertTo2D;
            bool processGrids = win.ProcessGrids;
            bool processLevels = win.ProcessLevels;
            List<int> selectedIndices = win.SelectedIndices;

            // ── Map indices back to View objects ──
            List<View> selectedViews = selectedIndices
                .Where(i => i >= 0 && i < allViews.Count)
                .Select(i => allViews[i])
                .ToList();

            // ── Target extent type ──
            DatumExtentType targetType = to2D
                ? DatumExtentType.ViewSpecific
                : DatumExtentType.Model;

            int gridCount = 0;
            int levelCount = 0;
            int viewsProcessed = 0;

            using (Transaction tx = new Transaction(doc, "Set Grid/Level Extent"))
            {
                tx.Start();

                foreach (View view in selectedViews)
                {
                    bool didWork = false;

                    if (processGrids)
                    {
                        List<Grid> grids = new FilteredElementCollector(doc, view.Id)
                            .OfClass(typeof(Grid))
                            .Cast<Grid>()
                            .ToList();

                        foreach (Grid g in grids)
                        {
                            try
                            {
                                g.SetDatumExtentType(
                                    DatumEnds.End0, view, targetType);
                                g.SetDatumExtentType(
                                    DatumEnds.End1, view, targetType);
                                gridCount++;
                                didWork = true;
                            }
                            catch { /* skip if not applicable */ }
                        }
                    }

                    if (processLevels)
                    {
                        List<Level> levels = new FilteredElementCollector(doc, view.Id)
                            .OfClass(typeof(Level))
                            .Cast<Level>()
                            .ToList();

                        foreach (Level lv in levels)
                        {
                            try
                            {
                                lv.SetDatumExtentType(
                                    DatumEnds.End0, view, targetType);
                                lv.SetDatumExtentType(
                                    DatumEnds.End1, view, targetType);
                                levelCount++;
                                didWork = true;
                            }
                            catch { /* skip if not applicable */ }
                        }
                    }

                    if (didWork) viewsProcessed++;
                }

                tx.Commit();
            }

            // ── Report ──
            string mode = to2D ? "2D (ViewSpecific)" : "3D (Model)";
            TaskDialog.Show("Grid & Level Extent \u2014 Done",
                $"Converted to: {mode}\n\n"
                + $"Grids processed:  {gridCount}\n"
                + $"Levels processed: {levelCount}\n"
                + $"Views affected:   {viewsProcessed}");

            return Result.Succeeded;
        }
    }
}