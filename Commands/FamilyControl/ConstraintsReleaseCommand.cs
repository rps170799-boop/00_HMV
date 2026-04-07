using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class ConstraintsReleaseCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // 1. Get the pre-selected elements
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

            if (selectedIds.Count == 0)
            {
                TaskDialog.Show("HMV Tools", "Please select one or more objects first.");
                return Result.Cancelled;
            }

            // 2. Open the Window UI
            var window = new ConstraintsReleaseWindow(selectedIds.Count);
            if (window.ShowDialog() != true)
            {
                return Result.Cancelled;
            }

            // 3. Process the total constraints release
            using (Transaction trans = new Transaction(doc, "Total Constraint Release"))
            {
                trans.Start();
                int constraintsRemoved = 0;

                foreach (ElementId id in selectedIds)
                {
                    Element el = doc.GetElement(id);
                    if (el == null) continue;

                    // A. UNPIN the element itself
                    if (el.Pinned)
                    {
                        el.Pinned = false;
                        constraintsRemoved++;
                    }

                    // B. Get ALL dependent elements (this catches constraints from surrounding items too)
                    ICollection<ElementId> dependentIds = el.GetDependentElements(null);

                    foreach (ElementId depId in dependentIds)
                    {
                        Element depEl = doc.GetElement(depId);
                        if (depEl == null) continue;

                        // C. Remove ALIGNMENT Constraints (the invisible blue padlocks)
                        if (depEl.Category != null && depEl.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Constraints)
                        {
                            try
                            {
                                doc.Delete(depId);
                                constraintsRemoved++;
                            }
                            catch { }
                            continue;
                        }

                        // D. Unlock DIMENSIONS, LABELS, and EQ CONSTRAINTS (without deleting the visual dimension)
                        if (depEl is Dimension dim)
                        {
                            bool changed = false;

                            // Unlock standard dimension lock
                            if (dim.IsLocked)
                            {
                                try { dim.IsLocked = false; changed = true; } catch { }
                            }

                            // Remove parameter labels (e.g. Width = 500)
                            if (dim.FamilyLabel != null)
                            {
                                try { dim.FamilyLabel = null; changed = true; } catch { }
                            }

                            // Turn off EQ constraints
                            try
                            {
                                if (dim.AreSegmentsEqual)
                                {
                                    dim.AreSegmentsEqual = false;
                                    changed = true;
                                }
                            }
                            catch { } // Fails safely if dimension doesn't have multiple segments

                            if (changed) constraintsRemoved++;
                        }
                    }
                }

                trans.Commit();
                TaskDialog.Show("Success", $"Totally unconstrained!\n\nReleased {constraintsRemoved} constraint(s) (Pins, Alignments, EQ, and Locked Dimensions) from the selection and surrounding items.");
            }

            return Result.Succeeded;
        }
    }
}