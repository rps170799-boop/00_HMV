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
    public class GenericAnnotationTagCommand : IExternalCommand
    {
        private const double MmToFeet = 1.0 / 304.8;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View view = doc.ActiveView;

            // ── 1. Pick linked elements ─────────────────────────
            IList<Reference> pickedRefs;
            try
            {
                pickedRefs = uidoc.Selection.PickObjects(
                    ObjectType.LinkedElement,
                    "Select elements in the linked model. Press Finish when done.");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }

            if (pickedRefs == null || pickedRefs.Count == 0)
                return Result.Cancelled;

            // ── 2. Resolve link + collect parameter names ───────
            var resolvedElements = new List<ResolvedLinkedElement>();
            var paramNameSet = new HashSet<string>();

            foreach (Reference r in pickedRefs)
            {
                RevitLinkInstance linkInst =
                    doc.GetElement(r) as RevitLinkInstance;
                if (linkInst == null) continue;

                Document linkDoc = linkInst.GetLinkDocument();
                if (linkDoc == null) continue;

                Element linkedElem =
                    linkDoc.GetElement(r.LinkedElementId);
                if (linkedElem == null) continue;

                Transform t = linkInst.GetTotalTransform();

                XYZ center = GetElementCenter(linkedElem, t);
                if (center == null) continue;

                resolvedElements.Add(new ResolvedLinkedElement
                {
                    Element = linkedElem,
                    HostCenter = center,
                    LinkTransform = t
                });

                // Collect INSTANCE parameter names
                foreach (Parameter p in linkedElem.Parameters)
                {
                    string name = p.Definition?.Name;
                    if (!string.IsNullOrEmpty(name))
                        paramNameSet.Add(name);
                }

                // Collect TYPE parameter names
                ElementId typeId = linkedElem.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                {
                    Element elemType = linkDoc.GetElement(typeId);
                    if (elemType != null)
                    {
                        foreach (Parameter p in elemType.Parameters)
                        {
                            string name = p.Definition?.Name;
                            if (!string.IsNullOrEmpty(name))
                                paramNameSet.Add(name);
                        }
                    }
                }
            }

            if (resolvedElements.Count == 0)
            {
                TaskDialog.Show("HMV Tools",
                    "No valid linked elements could be resolved.");
                return Result.Cancelled;
            }

            // ── 3. Collect TextNoteTypes in host ────────────────
            var textTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .OrderBy(t => t.Name)
                .Select(t => new TextTypeEntry
                {
                    TypeName = t.Name,
                    TypeIdInt = t.Id.IntegerValue
                })
                .ToList();

            if (textTypes.Count == 0)
            {
                TaskDialog.Show("HMV Tools",
                    "No TextNote types found in the project.");
                return Result.Cancelled;
            }

            var paramNames = paramNameSet.OrderBy(n => n).ToList();

            // ── 4. Show window ──────────────────────────────────
            var win = new GenericAnnotationTagWindow(
                textTypes, paramNames);

            if (win.ShowDialog() != true || win.Settings == null)
                return Result.Cancelled;

            GenericAnnotationTagSettings settings = win.Settings;

            // ── 5. Resolve chosen TextNoteType ──────────────────
            ElementId textTypeId =
                new ElementId(settings.TextTypeIdInt);
            TextNoteType textType =
                doc.GetElement(textTypeId) as TextNoteType;

            if (textType == null)
            {
                TaskDialog.Show("HMV Tools",
                    "Selected text type not found.");
                return Result.Cancelled;
            }

            // ── 6. Prepare placement ────────────────────────────
            double offX = settings.OffsetXMm * MmToFeet;
            double offY = settings.OffsetYMm * MmToFeet;
            string srcParamName = settings.ParameterName;
            bool hasOffset = Math.Abs(offX) > 1e-9
                          || Math.Abs(offY) > 1e-9;

            int placed = 0;
            int skipped = 0;
            var skippedReasons = new List<string>();

            // Get view plane Z for annotation placement
            double viewZ = 0;
            if (view is ViewPlan vp)
            {
                PlanViewRange pvr = vp.GetViewRange();
                Level cutLevel = doc.GetElement(
                    pvr.GetLevelId(PlanViewPlane.CutPlane)) as Level;
                double cutOffset =
                    pvr.GetOffset(PlanViewPlane.CutPlane);
                viewZ = cutLevel != null
                    ? cutLevel.Elevation + cutOffset
                    : 0;
            }

            using (Transaction tx = new Transaction(doc,
                "HMV – Generic Annotation Tag"))
            {
                tx.Start();

                foreach (var re in resolvedElements)
                {
                    // Read source parameter (instance → type)
                    Parameter srcParam =
                        re.Element.LookupParameter(srcParamName);

                    if (srcParam == null)
                    {
                        ElementId typeId = re.Element.GetTypeId();
                        if (typeId != ElementId.InvalidElementId)
                        {
                            Element elemType =
                                re.Element.Document.GetElement(typeId);
                            if (elemType != null)
                                srcParam = elemType.LookupParameter(
                                    srcParamName);
                        }
                    }

                    if (srcParam == null)
                    {
                        skipped++;
                        skippedReasons.Add(
                            $"{re.Element.Name} (Id {re.Element.Id.IntegerValue})"
                          + $" → '{srcParamName}' not found (omitted)");
                        continue;
                    }

                    string value = GetParameterValueAsString(srcParam);
                    if (string.IsNullOrEmpty(value))
                    {
                        skipped++;
                        skippedReasons.Add(
                            $"{re.Element.Name} (Id {re.Element.Id.IntegerValue})"
                          + $" → '{srcParamName}' is empty");
                        continue;
                    }

                    // Element center projected to view plane
                    XYZ origin = new XYZ(
                        re.HostCenter.X,
                        re.HostCenter.Y,
                        viewZ);

                    // TextNote placement = origin + offset
                    XYZ textPoint = new XYZ(
                        origin.X + offX,
                        origin.Y + offY,
                        viewZ);

                    // Create TextNote
                    TextNote note;
                    try
                    {
                        note = TextNote.Create(
                            doc, view.Id, textPoint,
                            value, textTypeId);
                    }
                    catch (Exception ex)
                    {
                        skipped++;
                        skippedReasons.Add(
                            $"{re.Element.Name} → TextNote error: {ex.Message}");
                        continue;
                    }

                    if (note == null)
                    {
                        skipped++;
                        skippedReasons.Add(
                            $"{re.Element.Name} → TextNote.Create returned null");
                        continue;
                    }

                    // Add real leader pointing to element center
                    if (hasOffset)
                    {
                        try
                        {
                            note.AddLeader(
                                TextNoteLeaderTypes.TNLT_STRAIGHT_L);
                            IList<Leader> leaders = note.GetLeaders();
                            if (leaders.Count > 0)
                                leaders[0].End = origin;
                        }
                        catch { /* leader failed — text still placed */ }
                    }

                    placed++;
                }

                tx.Commit();
            }

            // ── 7. Report ───────────────────────────────────────
            string report =
                "GENERIC ANNOTATION TAG SUMMARY\n"
              + "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
              + $"Elements picked:      {resolvedElements.Count}\n"
              + $"TextNotes placed:     {placed}\n"
              + $"Skipped:              {skipped}\n"
              + $"Text type:            {textType.Name}\n"
              + $"Source parameter:      {srcParamName}\n";

            if (hasOffset)
                report += $"Offset:               "
                        + $"X={settings.OffsetXMm}mm, "
                        + $"Y={settings.OffsetYMm}mm (with leader)\n";

            if (skippedReasons.Count > 0)
            {
                report += "\n── SKIPPED ──\n";
                foreach (string s in skippedReasons)
                    report += $"  • {s}\n";
            }

            var td = new TaskDialog("HMV Tools – Annotation Tag");
            td.MainContent = report;
            td.Show();

            return Result.Succeeded;
        }

        // ════════════════════════════════════════════════════════
        // Helpers
        // ════════════════════════════════════════════════════════

        private XYZ GetElementCenter(Element elem, Transform t)
        {
            Location loc = elem.Location;

            if (loc is LocationPoint lp)
                return t.OfPoint(lp.Point);

            if (loc is LocationCurve lc)
            {
                Curve c = lc.Curve;
                XYZ mid = c.Evaluate(0.5, true);
                return t.OfPoint(mid);
            }

            BoundingBoxXYZ bb = elem.get_BoundingBox(null);
            if (bb != null)
                return t.OfPoint((bb.Min + bb.Max) / 2.0);

            return null;
        }

        private string GetParameterValueAsString(Parameter p)
        {
            if (!p.HasValue) return null;

            switch (p.StorageType)
            {
                case StorageType.String:
                    return p.AsString();

                case StorageType.Integer:
                    return p.AsInteger().ToString();

                case StorageType.Double:
                    string vs = p.AsValueString();
                    return !string.IsNullOrEmpty(vs)
                        ? vs
                        : p.AsDouble().ToString("F3");

                case StorageType.ElementId:
                    ElementId id = p.AsElementId();
                    if (id == ElementId.InvalidElementId)
                        return null;
                    return p.AsValueString()
                        ?? id.IntegerValue.ToString();

                default:
                    return null;
            }
        }

        private class ResolvedLinkedElement
        {
            public Element Element;
            public XYZ HostCenter;
            public Transform LinkTransform;
        }
    }
}