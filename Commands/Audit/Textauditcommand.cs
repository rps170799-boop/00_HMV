using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class TextAuditCommand : IExternalCommand
    {
        // ── Standards (set from config window) ────────────────────
        private string _stdFont;
        private double _stdWidth;
        private string _stdNameTemplate;
        private const int STD_BOLD = 0;
        // Revit API: TEXT_BACKGROUND  0 = Opaque, 1 = Transparent
        private const int STD_OPAQUE = 0;

        private static readonly double[] STD_SIZES =
            { 1.5, 2.0, 2.5, 3.0, 3.5 };

        // ── Counters for unified report ───────────────────────────
        private int _typesStandardized;
        private int _instancesReassigned;
        private int _typesDeleted;
        private int _typesSkippedDeletion;
        private int _familiesChanged;
        private int _familiesCompliant;
        private List<string> _propertyChanges;
        private List<string> _mergeDetails;
        private List<string> _deleteDetails;
        private List<string> _familyDetails;
        private List<string> _allErrors;

        // ── Entry point ───────────────────────────────────────────

        public Result Execute(ExternalCommandData commandData,
            ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // ── Show config window ────────────────────────
                var config = new TextAuditConfigWindow();
                bool? result = config.ShowDialog();

                if (result != true)
                    return Result.Cancelled;

                _stdFont = config.SelectedFont;
                _stdWidth = config.SelectedWidth;
                _stdNameTemplate = config.SelectedNameTemplate;

                // Init counters
                _typesStandardized = 0;
                _instancesReassigned = 0;
                _typesDeleted = 0;
                _typesSkippedDeletion = 0;
                _familiesChanged = 0;
                _familiesCompliant = 0;
                _propertyChanges = new List<string>();
                _mergeDetails = new List<string>();
                _deleteDetails = new List<string>();
                _familyDetails = new List<string>();
                _allErrors = new List<string>();

                // ── STEP 1 ───────────────────────────────────
                Step01A_StandardizeAndGroup(doc,
                    out List<double> sizeList,
                    out List<List<int>> groupIdList,
                    out List<int> keeperIdList);

                // ── STEP 2 ───────────────────────────────────
                Step01B_ReassignInstances(doc,
                    sizeList, groupIdList, keeperIdList,
                    out List<int> typesToDelete);

                // ── STEP 3 ───────────────────────────────────
                Step01C_DeleteEmptyTypes(doc, typesToDelete);

                // ── STEP 4 ───────────────────────────────────
                Step02_StandardizeTagFamilies(doc);

                // ── Build unified report ─────────────────────
                string report = BuildReport();

                var window = new TextAuditReportWindow(report);
                window.ShowDialog();

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions
                       .OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // ═══════════════════════════════════════════════════════════
        //  UNIFIED REPORT
        // ═══════════════════════════════════════════════════════════

        private string BuildReport()
        {
            var sb = new StringBuilder();

            sb.AppendLine("TEXT AUDIT REPORT");
            sb.AppendLine(
                $"Font: {_stdFont}  |  Width: {_stdWidth}  |  " +
                $"Template: {_stdNameTemplate}");
            sb.AppendLine(new string('─', 52));

            // ── Summary ─────────────────────────────────────────
            sb.AppendLine(
                $"  Types standardized ......... {_typesStandardized}");
            sb.AppendLine(
                $"  Instances reassigned ....... {_instancesReassigned}");
            sb.AppendLine(
                $"  Duplicates deleted ......... {_typesDeleted}");
            sb.AppendLine(
                $"  Skipped (still in use) ..... {_typesSkippedDeletion}");
            sb.AppendLine(
                $"  Tag families updated ....... {_familiesChanged}");
            sb.AppendLine(
                $"  Tag families compliant ..... {_familiesCompliant}");

            // ── Property changes ────────────────────────────────
            if (_propertyChanges.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("PROPERTY CHANGES");
                sb.AppendLine(new string('─', 52));
                foreach (string c in _propertyChanges)
                    sb.AppendLine(c);
            }

            // ── Merge / reassign details ────────────────────────
            if (_mergeDetails.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("MERGE & REASSIGNMENT");
                sb.AppendLine(new string('─', 52));
                foreach (string m in _mergeDetails)
                    sb.AppendLine(m);
            }

            // ── Deletion details ────────────────────────────────
            if (_deleteDetails.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("PURGED TYPES");
                sb.AppendLine(new string('─', 52));
                foreach (string d in _deleteDetails)
                    sb.AppendLine(d);
            }

            // ── Tag family details ──────────────────────────────
            if (_familyDetails.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("TAG FAMILIES UPDATED");
                sb.AppendLine(new string('─', 52));
                foreach (string f in _familyDetails)
                    sb.AppendLine(f);
            }

            // ── Errors ──────────────────────────────────────────
            if (_allErrors.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine($"ERRORS ({_allErrors.Count})");
                sb.AppendLine(new string('─', 52));
                foreach (string e in _allErrors)
                    sb.AppendLine("  - " + e);
            }

            if (_allErrors.Count == 0)
            {
                sb.AppendLine();
                sb.AppendLine("✓ No errors.");
            }

            return sb.ToString();
        }

        // ═══════════════════════════════════════════════════════════
        //  STEP 01A – Standardize properties + group by size
        // ═══════════════════════════════════════════════════════════

        private void Step01A_StandardizeAndGroup(Document doc,
            out List<double> sizeList,
            out List<List<int>> groupIdList,
            out List<int> keeperIdList)
        {
            // ── Collect types ──────────────────────────────────
            var allTypes = new FilteredElementCollector(doc)
                .WhereElementIsElementType()
                .ToElements();

            var textNoteTypes = new List<Element>();

            foreach (Element t in allTypes)
            {
                Parameter pFont = t.get_Parameter(
                    BuiltInParameter.TEXT_FONT);
                if (pFont == null) continue;

                string className = t.GetType().Name;
                if (className == "TextNoteType")
                    textNoteTypes.Add(t);
            }

            // ── Count instances per type ───────────────────────
            var instanceCount = new Dictionary<int, int>();

            foreach (TextNote inst in new FilteredElementCollector(doc)
                .OfClass(typeof(TextNote))
                .WhereElementIsNotElementType())
            {
                int tid = inst.GetTypeId().IntegerValue;
                if (!instanceCount.ContainsKey(tid))
                    instanceCount[tid] = 0;
                instanceCount[tid]++;
            }

            // ── Standardize properties (Transaction) ───────────
            using (var tx = new Transaction(doc,
                "HMV Text Audit – Standardize Properties"))
            {
                tx.Start();

                var allTargets = textNoteTypes.ToList();

                foreach (Element t in allTargets)
                {
                    string name = t.Name;
                    int eid = t.Id.IntegerValue;
                    string cls = t.GetType().Name;
                    var props = new List<string>();

                    TrySet_Font(t, props);
                    TrySet_Width(t, props);
                    TrySet_Bold(t, props);
                    TrySet_Background(t, props);
                    TrySet_Size(t, props);

                    if (props.Count > 0)
                    {
                        _typesStandardized++;
                        _propertyChanges.Add(string.Format(
                            "{0} [{1}] (ID: {2})\n   {3}",
                            name, cls, eid,
                            string.Join("\n   ", props)));
                    }
                }

                tx.Commit();
            }

            // ── Group TextNoteTypes by size ────────────────────
            var sizeGroups = new Dictionary<double, List<GroupItem>>();
            foreach (double s in STD_SIZES)
                sizeGroups[s] = new List<GroupItem>();

            foreach (Element t in textNoteTypes)
            {
                Parameter p = t.get_Parameter(
                    BuiltInParameter.TEXT_SIZE);
                if (p == null) continue;

                double sizeMm = Math.Round(
                    p.AsDouble() * 304.8, 2);
                double target = NearestSize(sizeMm);
                int count = 0;
                instanceCount.TryGetValue(
                    t.Id.IntegerValue, out count);

                sizeGroups[target].Add(new GroupItem
                {
                    Id = t.Id.IntegerValue,
                    Name = t.Name,
                    Count = count
                });
            }

            // ── Find keepers per size ──────────────────────────
            var keeperPerSize = new Dictionary<double, GroupItem>();

            foreach (double size in STD_SIZES)
            {
                string stdName = string.Format(
                    _stdNameTemplate, size);
                var group = sizeGroups[size];

                if (group.Count == 0)
                {
                    keeperPerSize[size] = null;
                    continue;
                }

                GroupItem keeper = group.FirstOrDefault(
                    g => g.Name == stdName);

                if (keeper == null)
                    keeper = group.OrderByDescending(
                        g => g.Count).First();

                keeperPerSize[size] = keeper;
            }

            // ── Output lists ───────────────────────────────────
            sizeList = new List<double>(STD_SIZES);
            groupIdList = new List<List<int>>();
            keeperIdList = new List<int>();

            foreach (double size in STD_SIZES)
            {
                groupIdList.Add(
                    sizeGroups[size].Select(g => g.Id).ToList());
                keeperIdList.Add(
                    keeperPerSize[size]?.Id ?? -1);
            }
        }

        // ═══════════════════════════════════════════════════════════
        //  STEP 01B – Reassign instances to keepers
        // ═══════════════════════════════════════════════════════════

        private void Step01B_ReassignInstances(Document doc,
            List<double> sizes,
            List<List<int>> groupIds,
            List<int> keeperIds,
            out List<int> typesToDelete)
        {
            typesToDelete = new List<int>();

            // Map type → instances
            var typeToInstances =
                new Dictionary<int, List<TextNote>>();

            foreach (TextNote inst in
                new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNote))
                    .WhereElementIsNotElementType())
            {
                int tid = inst.GetTypeId().IntegerValue;
                if (!typeToInstances.ContainsKey(tid))
                    typeToInstances[tid] = new List<TextNote>();
                typeToInstances[tid].Add(inst);
            }

            using (var tx = new Transaction(doc,
                "HMV Text Audit – Reassign Instances"))
            {
                tx.Start();

                for (int i = 0; i < sizes.Count; i++)
                {
                    double size = sizes[i];
                    var group = groupIds[i];
                    int keeperIdInt = keeperIds[i];
                    string stdName = string.Format(
                        _stdNameTemplate, size);

                    if (keeperIdInt == -1 || group.Count == 0)
                        continue;

                    // Find keeper element
                    Element keeperElem = null;
                    foreach (int eid in group)
                    {
                        Element e = doc.GetElement(
                            new ElementId(eid));
                        if (e != null &&
                            e.Id.IntegerValue == keeperIdInt)
                        {
                            keeperElem = e;
                            break;
                        }
                    }

                    if (keeperElem == null)
                    {
                        _allErrors.Add(
                            $"{size}mm: Keeper ID {keeperIdInt}" +
                            " not found");
                        continue;
                    }

                    // Check if standard name already exists
                    Element stdElem = null;
                    foreach (int eid in group)
                    {
                        Element e = doc.GetElement(
                            new ElementId(eid));
                        if (e != null && e.Name == stdName)
                        {
                            stdElem = e;
                            break;
                        }
                    }

                    Element finalKeeper;

                    if (stdElem != null)
                    {
                        finalKeeper = stdElem;
                    }
                    else
                    {
                        try
                        {
                            keeperElem.Name = stdName;
                            finalKeeper = keeperElem;
                        }
                        catch
                        {
                            try
                            {
                                var tntKeeper =
                                    keeperElem as TextNoteType;
                                if (tntKeeper != null)
                                {
                                    var dup = tntKeeper.Duplicate(
                                        stdName);
                                    finalKeeper = dup;
                                }
                                else
                                {
                                    _allErrors.Add(
                                        $"{size}mm: Could not " +
                                        "rename or duplicate");
                                    continue;
                                }
                            }
                            catch (Exception ex)
                            {
                                _allErrors.Add(
                                    $"{size}mm: {ex.Message}");
                                continue;
                            }
                        }
                    }

                    ElementId finalKeeperId = finalKeeper.Id;
                    int reassigned = 0;
                    int merged = 0;

                    foreach (int eid in group)
                    {
                        if (eid == finalKeeperId.IntegerValue)
                            continue;

                        List<TextNote> instances;
                        if (!typeToInstances.TryGetValue(
                            eid, out instances))
                            instances = new List<TextNote>();

                        foreach (TextNote inst in instances)
                        {
                            try
                            {
                                inst.ChangeTypeId(finalKeeperId);
                                reassigned++;
                            }
                            catch (Exception ex)
                            {
                                _allErrors.Add(
                                    $"Instance " +
                                    $"{inst.Id.IntegerValue}: " +
                                    $"{ex.Message}");
                            }
                        }

                        merged++;
                        typesToDelete.Add(eid);
                    }

                    _instancesReassigned += reassigned;

                    if (merged > 0 || reassigned > 0)
                    {
                        _mergeDetails.Add(
                            $"  {size}mm -> \"{stdName}\": " +
                            $"{merged} merged, " +
                            $"{reassigned} reassigned");
                    }
                }

                tx.Commit();
            }
        }

        // ═══════════════════════════════════════════════════════════
        //  STEP 01C – Delete empty types
        // ═══════════════════════════════════════════════════════════

        private void Step01C_DeleteEmptyTypes(Document doc,
            List<int> typesToDelete)
        {
            if (typesToDelete == null || typesToDelete.Count == 0)
                return;

            // Recount instances
            var typeCount = new Dictionary<int, int>();

            foreach (TextNote inst in
                new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNote))
                    .WhereElementIsNotElementType())
            {
                int tid = inst.GetTypeId().IntegerValue;
                if (!typeCount.ContainsKey(tid))
                    typeCount[tid] = 0;
                typeCount[tid]++;
            }

            using (var tx = new Transaction(doc,
                "HMV Text Audit – Delete Empty Types"))
            {
                tx.Start();

                foreach (int eid in typesToDelete)
                {
                    Element elem = doc.GetElement(
                        new ElementId(eid));
                    if (elem == null)
                    {
                        _allErrors.Add(
                            $"ID {eid}: element not found");
                        continue;
                    }

                    string name = elem.Name;
                    int count = 0;
                    typeCount.TryGetValue(eid, out count);

                    if (count > 0)
                    {
                        _typesSkippedDeletion++;
                        _deleteDetails.Add(
                            $"  SKIP: {name} " +
                            $"({count} instances remain)");
                        continue;
                    }

                    try
                    {
                        doc.Delete(elem.Id);
                        _typesDeleted++;
                        _deleteDetails.Add(
                            $"  DEL:  {name}");
                    }
                    catch (Exception ex)
                    {
                        _allErrors.Add(
                            $"{name} (ID: {eid}): " +
                            ex.Message);
                    }
                }

                tx.Commit();
            }
        }

        // ═══════════════════════════════════════════════════════════
        //  STEP 02 – Standardize annotation tag families
        // ═══════════════════════════════════════════════════════════

        private void Step02_StandardizeTagFamilies(Document doc)
        {
            var annotFams = new List<Family>();

            foreach (Family f in
                new FilteredElementCollector(doc)
                    .OfClass(typeof(Family)))
            {
                try
                {
                    if (f.FamilyCategory != null &&
                        f.FamilyCategory.CategoryType ==
                            CategoryType.Annotation &&
                        !f.IsInPlace && f.IsEditable)
                    {
                        annotFams.Add(f);
                    }
                }
                catch { /* skip */ }
            }

            var loadOpts = new TextAuditFamilyLoadOptions();

            foreach (Family f in annotFams)
            {
                string famName = f.Name;

                if (famName.Contains("Membrete") ||
                    famName.Contains("Pagina Inicio"))
                {
                    _familiesCompliant++;
                    continue;
                }

                Document fDoc = null;
                try
                {
                    fDoc = doc.EditFamily(f);
                    if (fDoc == null)
                    {
                        _familiesCompliant++;
                        continue;
                    }

                    var famTypes =
                        new FilteredElementCollector(fDoc)
                            .WhereElementIsElementType()
                            .ToElements();
                    var famInstances =
                        new FilteredElementCollector(fDoc)
                            .WhereElementIsNotElementType()
                            .ToElements();

                    var famAll = famTypes.Concat(famInstances)
                        .ToList();

                    bool hasChanges = false;
                    int elemCount = 0;

                    using (var t = new Transaction(fDoc,
                        "Standardize All Properties"))
                    {
                        t.Start();

                        foreach (Element et in famAll)
                        {
                            var typeChanges = new List<string>();

                            TrySet_Font(et, typeChanges);
                            TrySet_Bold(et, typeChanges);
                            TrySet_Width(et, typeChanges);
                            TrySet_Background(et, typeChanges);
                            TrySet_Size(et, typeChanges);

                            if (typeChanges.Count > 0)
                            {
                                hasChanges = true;
                                elemCount++;
                            }
                        }

                        t.Commit();
                    }

                    if (hasChanges)
                    {
                        try
                        {
                            fDoc.LoadFamily(doc, loadOpts);
                            _familiesChanged++;
                            _familyDetails.Add(
                                $"  {famName}: " +
                                $"{elemCount} elements changed");
                        }
                        catch (Exception ex)
                        {
                            _allErrors.Add(
                                $"{famName}: reload failed" +
                                $" - {ex.Message}");
                        }
                    }
                    else
                    {
                        _familiesCompliant++;
                    }

                    fDoc.Close(false);
                }
                catch (Exception ex)
                {
                    _allErrors.Add($"{famName}: {ex.Message}");
                    if (fDoc != null)
                    {
                        try { fDoc.Close(false); }
                        catch { /* ignore */ }
                    }
                }
            }
        }

        // ═══════════════════════════════════════════════════════════
        //  Helpers
        // ═══════════════════════════════════════════════════════════

        private static double NearestSize(double valMm)
        {
            double closest = STD_SIZES[0];
            double minDiff = Math.Abs(valMm - closest);

            for (int i = 1; i < STD_SIZES.Length; i++)
            {
                double diff = Math.Abs(valMm - STD_SIZES[i]);
                if (diff < minDiff)
                {
                    minDiff = diff;
                    closest = STD_SIZES[i];
                }
            }
            return closest;
        }

        // ── Parameter setters ──────────────────────────────────

        private void TrySet_Font(Element t,
            List<string> props)
        {
            try
            {
                Parameter p = t.get_Parameter(
                    BuiltInParameter.TEXT_FONT);
                if (p != null && !p.IsReadOnly &&
                    p.AsString() != _stdFont)
                {
                    string old = p.AsString();
                    p.Set(_stdFont);
                    props.Add($"Font: {old} -> {_stdFont}");
                }
            }
            catch { /* skip */ }
        }

        private void TrySet_Width(Element t,
            List<string> props)
        {
            try
            {
                Parameter p = t.get_Parameter(
                    BuiltInParameter.TEXT_WIDTH_SCALE);
                if (p != null && !p.IsReadOnly &&
                    Math.Round(p.AsDouble(), 4) != _stdWidth)
                {
                    double old = Math.Round(p.AsDouble(), 4);
                    p.Set(_stdWidth);
                    props.Add($"Width: {old} -> {_stdWidth}");
                }
            }
            catch { /* skip */ }
        }

        private static void TrySet_Bold(Element t,
            List<string> props)
        {
            try
            {
                Parameter p = t.get_Parameter(
                    BuiltInParameter.TEXT_STYLE_BOLD);
                if (p != null && !p.IsReadOnly &&
                    p.AsInteger() != STD_BOLD)
                {
                    string old = p.AsInteger() == 1
                        ? "Yes" : "No";
                    p.Set(STD_BOLD);
                    props.Add($"Bold: {old} -> No");
                }
            }
            catch { /* skip */ }
        }

        private static void TrySet_Background(Element t,
            List<string> props)
        {
            try
            {
                Parameter p = t.get_Parameter(
                    BuiltInParameter.TEXT_BACKGROUND);
                if (p != null && !p.IsReadOnly &&
                    p.AsInteger() != STD_OPAQUE)
                {
                    // Revit API: 0 = Opaque, 1 = Transparent
                    string old = p.AsInteger() == 0
                        ? "Opaque" : "Transparent";
                    p.Set(STD_OPAQUE);
                    props.Add($"Background: {old} -> Opaque");
                }
            }
            catch { /* skip */ }
        }

        private static void TrySet_Size(Element t,
            List<string> props)
        {
            try
            {
                Parameter p = t.get_Parameter(
                    BuiltInParameter.TEXT_SIZE);
                if (p != null && !p.IsReadOnly)
                {
                    double currentMm = Math.Round(
                        p.AsDouble() * 304.8, 4);
                    double targetMm = NearestSize(currentMm);

                    if (Math.Round(currentMm, 2) != targetMm)
                    {
                        p.Set(targetMm / 304.8);
                        props.Add(
                            $"Size: " +
                            $"{Math.Round(currentMm, 2)}mm " +
                            $"-> {targetMm}mm");
                    }
                }
            }
            catch { /* skip */ }
        }

        // ── Internal model ─────────────────────────────────────

        private class GroupItem
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public int Count { get; set; }
        }
    }
}