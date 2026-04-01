using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class SignageHostingCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // ==========================================
                // STEP 1: BATCH PICK PEDESTALS (Compound or Single)
                // ==========================================
                TaskDialog.Show("HMV Tools", "Step 1: Window-select ALL the PEDESTALS (You can drag a box over them). Click 'Finish' in the top left options bar when done.");

                IList<Reference> pedRefs;
                try
                {
                    ISelectionFilter pedFilter = new CategorySelectionFilter(BuiltInCategory.OST_StructuralFraming);
                    pedRefs = uidoc.Selection.PickObjects(ObjectType.Element, pedFilter, "Window-select all Pedestals");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                if (!pedRefs.Any()) return Result.Failed;

                // ==========================================
                // STEP 2: BATCH PICK EQUIPMENT
                // ==========================================
                TaskDialog.Show("HMV Tools", "Step 2: Window-select ALL the ELECTRICAL EQUIPMENT. Click 'Finish' when done.");

                IList<Reference> eqRefs;
                try
                {
                    ISelectionFilter eqFilter = new CategorySelectionFilter(BuiltInCategory.OST_ElectricalEquipment);
                    eqRefs = uidoc.Selection.PickObjects(ObjectType.Element, eqFilter, "Window-select all Electrical Equipment");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                if (!eqRefs.Any()) return Result.Failed;

                // ==========================================
                // STEP 3: GATHER UI DATA (WITH RECURSIVE SEARCH)
                // ==========================================
                FamilyInstance firstPedestal = doc.GetElement(pedRefs.First()) as FamilyInstance;
                FamilyInstance firstEquipment = doc.GetElement(eqRefs.First()) as FamilyInstance;

                var signageCollector = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Signage)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .ToList();

                if (!signageCollector.Any())
                {
                    TaskDialog.Show("Error", "No Signage families found! Check if they are categorized as OST_Signage.");
                    return Result.Failed;
                }
                List<string> signageTypeNames = signageCollector.Select(s => $"{s.FamilyName} : {s.Name}").OrderBy(n => n).ToList();

                // 🌟 RECURSIVE DEEP SEARCH FOR UI DROPDOWN 🌟
                HashSet<string> nestedNamesSet = new HashSet<string>();
                GetNestedNamesRecursive(doc, firstPedestal, nestedNamesSet);
                List<string> nestedNames = nestedNamesSet.OrderBy(n => n).ToList();

                List<string> equipmentParams = firstEquipment.Parameters
                    .Cast<Parameter>()
                    .Where(p => p.StorageType != StorageType.ElementId && p.StorageType != StorageType.None)
                    .Select(p => p.Definition.Name)
                    .Distinct()
                    .OrderBy(n => n).ToList();

                // HARVEST SIGNAGE PARAMETERS
                List<string> targetParams = new List<string>();
                var existingSignages = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Signage)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>();

                foreach (var sig in existingSignages)
                {
                    foreach (Parameter p in sig.Parameters)
                    {
                        if (!p.IsReadOnly && p.StorageType != StorageType.ElementId && p.StorageType != StorageType.None)
                            targetParams.Add(p.Definition.Name);
                    }
                }

                if (!targetParams.Any())
                {
                    using (Transaction t = new Transaction(doc, "DummyExtraction"))
                    {
                        t.Start();
                        FailureHandlingOptions failOpt = t.GetFailureHandlingOptions();
                        failOpt.SetFailuresPreprocessor(new WarningSuppressor());
                        t.SetFailureHandlingOptions(failOpt);

                        foreach (var symbol in signageCollector)
                        {
                            if (!symbol.IsActive) symbol.Activate();
                            try
                            {
                                FamilyInstance dummy = doc.Create.NewFamilyInstance(XYZ.Zero, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                if (dummy != null)
                                {
                                    foreach (Parameter p in dummy.Parameters)
                                    {
                                        if (!p.IsReadOnly && p.StorageType != StorageType.ElementId && p.StorageType != StorageType.None)
                                            targetParams.Add(p.Definition.Name);
                                    }
                                }
                            }
                            catch { /* Ignore */ }
                        }
                        t.RollBack();
                    }
                }

                targetParams.AddRange(new[] { "Comments", "Mark" });
                targetParams = targetParams.Distinct().OrderBy(n => n).ToList();

                // ==========================================
                // STEP 4: SHOW UI WINDOW
                // ==========================================
                var window = new SignageSettingsWindow(signageTypeNames, nestedNames, equipmentParams, targetParams);
                if (window.ShowDialog() != true) return Result.Cancelled;

                string selectedSignageString = window.SelectedSignageType;
                string selectedNestedName = window.SelectedNestedFamily;
                string sourceParamName = window.SelectedSourceParam;
                string targetParamName = window.SelectedTargetParam;

                string[] parts = selectedSignageString.Split(new[] { " : " }, StringSplitOptions.None);
                FamilySymbol signageSymbol = signageCollector.FirstOrDefault(s => s.FamilyName == parts[0] && s.Name == parts[1]);
                if (signageSymbol == null) return Result.Failed;

                // ==========================================
                // STEP 5: DEEP NESTED BATCH PROCESSING LOOP
                // ==========================================
                int successCount = 0;

                using (Transaction tx = new Transaction(doc, "Batch Host Signage"))
                {
                    tx.Start();
                    if (!signageSymbol.IsActive) signageSymbol.Activate();

                    foreach (Reference pedRef in pedRefs)
                    {
                        FamilyInstance currentPedestal = doc.GetElement(pedRef) as FamilyInstance;
                        if (currentPedestal == null) continue;

                        // 🌟 FIND ALL TARGET ANCHORS, NO MATTER HOW DEEPLY NESTED 🌟
                        List<FamilyInstance> targetAnchors = new List<FamilyInstance>();
                        GetTargetAnchorsRecursive(doc, currentPedestal, selectedNestedName, targetAnchors);

                        // If no anchors found (fallback to standard origin placement)
                        if (targetAnchors.Count == 0)
                        {
                            targetAnchors.Add(currentPedestal); // Treat the main pedestal as the only anchor
                        }

                        // LOOP THROUGH EVERY SINGLE ANCHOR FOUND IN THIS FAMILY
                        foreach (FamilyInstance anchor in targetAnchors)
                        {
                            // --- 5A: FIND CLOSEST EQUIPMENT TO THIS SPECIFIC ANCHOR ---
                            XYZ anchorCenter = GetElementCenter(anchor);
                            FamilyInstance closestEquipment = null;
                            double minDistance = double.MaxValue;

                            foreach (Reference eqRef in eqRefs)
                            {
                                FamilyInstance eq = doc.GetElement(eqRef) as FamilyInstance;
                                if (eq == null) continue;

                                double dist = anchorCenter.DistanceTo(GetElementCenter(eq));
                                if (dist < minDistance)
                                {
                                    minDistance = dist;
                                    closestEquipment = eq;
                                }
                            }

                            if (closestEquipment == null) continue;

                            // --- 5B: PLACE & ROTATE ---
                            try
                            {
                                XYZ initialPlacementPoint = anchor.GetTransform().Origin;
                                double rotationAngle = XYZ.BasisX.AngleOnPlaneTo(anchor.GetTransform().BasisX, XYZ.BasisZ);

                                Level pedLevel = doc.GetElement(currentPedestal.LevelId) as Level;
                                FamilyInstance newSignage;

                                if (pedLevel != null)
                                    newSignage = doc.Create.NewFamilyInstance(initialPlacementPoint, signageSymbol, pedLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                else
                                    newSignage = doc.Create.NewFamilyInstance(initialPlacementPoint, signageSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                if (Math.Abs(rotationAngle) > 0.0001)
                                {
                                    Transform sigTransform = newSignage.GetTransform();
                                    double currentAngle = XYZ.BasisX.AngleOnPlaneTo(sigTransform.BasisX, XYZ.BasisZ);
                                    double angleDiff = rotationAngle - currentAngle;

                                    if (Math.Abs(angleDiff) > 0.0001)
                                    {
                                        Line axis = Line.CreateBound(initialPlacementPoint, initialPlacementPoint + XYZ.BasisZ);
                                        ElementTransformUtils.RotateElement(doc, newSignage.Id, axis, angleDiff);
                                    }
                                }

                                doc.Regenerate(); // Crucial for bounding box update

                                // --- 5C: ALIGN PHYSICAL CENTERS TO FRONT FACE ---
                                if (anchor.Name == selectedNestedName) // Only align if it's the actual anchor plate
                                {
                                    BoundingBoxXYZ anchorBox = anchor.get_BoundingBox(null);
                                    BoundingBoxXYZ signBox = newSignage.get_BoundingBox(null);

                                    if (anchorBox != null && signBox != null)
                                    {
                                        XYZ physicalAnchorCenter = (anchorBox.Min + anchorBox.Max) / 2.0;
                                        XYZ signCenter = (signBox.Min + signBox.Max) / 2.0;

                                        Transform anchorTx = anchor.GetTransform();

                                        XYZ localX = anchorTx.BasisX.Normalize();
                                        XYZ localY = anchor.FacingOrientation.Normalize();
                                        if (localY.IsAlmostEqualTo(XYZ.Zero)) localY = anchorTx.BasisY.Normalize();
                                        XYZ localZ = localX.CrossProduct(localY).Normalize();

                                        // 1. Find the raw 3D distance
                                        XYZ centerDiff = physicalAnchorCenter - signCenter;

                                        // 2. Project this distance onto the Plate's local face
                                        double alignX = centerDiff.DotProduct(localX);
                                        double alignZ = centerDiff.DotProduct(localZ);

                                        // 🌟 THIS IS WHAT WAS MISSING! Distance to center it on the Y depth 🌟
                                        double alignY = centerDiff.DotProduct(localY);

                                        // 3. Calculate the thickness depth
                                        XYZ anchorDiag = anchorBox.Max - anchorBox.Min;
                                        XYZ signDiag = signBox.Max - signBox.Min;

                                        double anchorThickness = Math.Abs(anchorDiag.DotProduct(localY));
                                        double signThickness = Math.Abs(signDiag.DotProduct(localY));

                                        double pushOutwardFeet = (anchorThickness / 2.0) + (signThickness / 2.0) + (2.0 / 304.8);

                                        // 4. Combine it all
                                        // Center it (alignY) THEN push it to the front face (pushOutwardFeet)
                                        double finalYMove = alignY - pushOutwardFeet;

                                        XYZ moveVector = (localX * alignX) + (localZ * alignZ) + (localY * finalYMove);

                                        if (!moveVector.IsAlmostEqualTo(XYZ.Zero))
                                        {
                                            ElementTransformUtils.MoveElement(doc, newSignage.Id, moveVector);
                                        }
                                    }
                                }

                                // --- 5D: SYNC PARAMETERS ---
                                Parameter srcParam = closestEquipment.LookupParameter(sourceParamName);
                                Parameter tgtParam = newSignage.LookupParameter(targetParamName);

                                if (srcParam != null && tgtParam != null && !tgtParam.IsReadOnly)
                                {
                                    switch (srcParam.StorageType)
                                    {
                                        case StorageType.String: tgtParam.Set(srcParam.AsString() ?? ""); break;
                                        case StorageType.Double: tgtParam.Set(srcParam.AsDouble()); break;
                                        case StorageType.Integer: tgtParam.Set(srcParam.AsInteger()); break;
                                    }
                                }

                                successCount++;
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }

                    tx.Commit();
                }

                TaskDialog.Show("Success", $"Successfully placed and aligned {successCount} signs in batch mode!");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("FATAL ERROR", ex.Message);
                return Result.Failed;
            }
        }

        // =========================================================================================
        // HELPER METHODS
        // =========================================================================================

        private XYZ GetElementCenter(Element el)
        {
            BoundingBoxXYZ box = el.get_BoundingBox(null);
            if (box == null) return (el.Location as LocationPoint)?.Point ?? XYZ.Zero;
            return (box.Min + box.Max) / 2.0;
        }

        // 🌟 RECURSIVE FUNCTION: Digs infinitely deep to find all nested family names 🌟
        private void GetNestedNamesRecursive(Document doc, FamilyInstance parent, HashSet<string> names)
        {
            foreach (ElementId id in parent.GetSubComponentIds())
            {
                if (doc.GetElement(id) is FamilyInstance subInst)
                {
                    names.Add(subInst.Name);
                    GetNestedNamesRecursive(doc, subInst, names); // Dig deeper
                }
            }
        }

        // 🌟 RECURSIVE FUNCTION: Digs infinitely deep to retrieve the actual instances 🌟
        private void GetTargetAnchorsRecursive(Document doc, FamilyInstance parent, string targetName, List<FamilyInstance> foundAnchors)
        {
            foreach (ElementId id in parent.GetSubComponentIds())
            {
                if (doc.GetElement(id) is FamilyInstance subInst)
                {
                    if (subInst.Name == targetName)
                    {
                        foundAnchors.Add(subInst);
                    }
                    GetTargetAnchorsRecursive(doc, subInst, targetName, foundAnchors); // Dig deeper
                }
            }
        }
    }

    public class CategorySelectionFilter : ISelectionFilter
    {
        private BuiltInCategory _cat;
        public CategorySelectionFilter(BuiltInCategory cat) { _cat = cat; }
        public bool AllowElement(Element elem) => elem.Category != null && elem.Category.Id.IntegerValue == (int)_cat;
        public bool AllowReference(Reference reference, XYZ position) => false;
    }

    public class WarningSuppressor : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            var failures = failuresAccessor.GetFailureMessages();
            foreach (var f in failures)
            {
                failuresAccessor.DeleteWarning(f);
            }
            return FailureProcessingResult.Continue;
        }
    }
}