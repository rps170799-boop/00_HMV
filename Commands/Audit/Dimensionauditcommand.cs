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
    public class DimensionAuditCommand : IExternalCommand
    {
        // ── Standards ──────────────────────────────────────────────
        private const string STD_FONT = "Arial";
        private const double STD_WIDTH = 1.0;
        private const int STD_BOLD = 0;
        // Revit API: TEXT_BACKGROUND  0 = Opaque, 1 = Transparent
        private const int STD_OPAQUE = 0;

        private static readonly double[] STD_SIZES =
            { 1.5, 2.0, 2.5, 3.0, 3.5 };

        // Name template:
        // HMV_Acotado Lineal {unit}_{size}mm Arial CIV.{XX|XXX}
        private const string STD_NAME_TEMPLATE =
            "HMV_Acotado Lineal {0}_{1}mm Arial CIV.{2}";

        // ── Counters for unified report ────────────────────────────
        private int _typesStandardized;
        private int _instancesReassigned;
        private int _typesDeleted;
        private int _typesSkippedDeletion;
        private int _instancesSkippedGroup;
        private List<string> _propertyChanges;
        private List<string> _mergeDetails;
        private List<string> _deleteDetails;
        private List<string> _allErrors;

        // ── Entry point ────────────────────────────────────────────

        public Result Execute(ExternalCommandData commandData,
            ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Init counters
                _typesStandardized = 0;
                _instancesReassigned = 0;
                _typesDeleted = 0;
                _typesSkippedDeletion = 0;
                _instancesSkippedGroup = 0;
                _propertyChanges = new List<string>();
                _mergeDetails = new List<string>();
                _deleteDetails = new List<string>();
                _allErrors = new List<string>();

                // ── STEP A ──────────────────────────────────────
                StepA_StandardizeAndGroup(doc,
                    out List<GroupKey> groupKeys,
                    out Dictionary<string, List<GroupItem>> groups,
                    out Dictionary<string, GroupItem> keepers);

                // ── STEP B ──────────────────────────────────────
                StepB_ReassignInstances(doc,
                    groupKeys, groups, keepers,
                    out List<int> typesToDelete);

                // ── STEP C ──────────────────────────────────────
                StepC_DeleteEmptyTypes(doc, typesToDelete);

                // ── Build unified report ────────────────────────
                string report = BuildReport();

                var window = new DimensionAuditReportWindow(report);
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

        // ════════════════════════════════════════════════════════════
        //  UNIFIED REPORT
        // ════════════════════════════════════════════════════════════

        private string BuildReport()
        {
            var sb = new StringBuilder();

            sb.AppendLine("HMV DIMENSION AUDIT - STANDARDIZATION REPORT");
            sb.AppendLine(new string('=', 55));
            sb.AppendLine();

            // ── Summary counts ─────────────────────────────────────
            sb.AppendLine("SUMMARY");
            sb.AppendLine(new string('-', 55));
            sb.AppendLine(
                $"  Types standardized (properties):  {_typesStandardized}");
            sb.AppendLine(
                $"  Instances reassigned to keepers:  {_instancesReassigned}");
            sb.AppendLine(
                $"  Instances skipped (in groups):    {_instancesSkippedGroup}");
            sb.AppendLine(
                $"  Duplicate types deleted:          {_typesDeleted}");
            sb.AppendLine(
                $"  Types skipped (still in use):     {_typesSkippedDeletion}");
            sb.AppendLine();

            // ── Property changes ───────────────────────────────────
            if (_propertyChanges.Count > 0)
            {
                sb.AppendLine("PROPERTY CHANGES");
                sb.AppendLine(new string('-', 55));
                foreach (string c in _propertyChanges)
                {
                    sb.AppendLine(c);
                    sb.AppendLine();
                }
            }

            // ── Merge / reassign details ───────────────────────────
            if (_mergeDetails.Count > 0)
            {
                sb.AppendLine("TYPE MERGE & REASSIGNMENT");
                sb.AppendLine(new string('-', 55));
                foreach (string m in _mergeDetails)
                    sb.AppendLine(m);
                sb.AppendLine();
            }

            // ── Deletion details ───────────────────────────────────
            if (_deleteDetails.Count > 0)
            {
                sb.AppendLine("PURGED TYPES");
                sb.AppendLine(new string('-', 55));
                foreach (string d in _deleteDetails)
                    sb.AppendLine(d);
                sb.AppendLine();
            }

            // ── Errors ─────────────────────────────────────────────
            if (_allErrors.Count > 0)
            {
                sb.AppendLine($"ERRORS ({_allErrors.Count})");
                sb.AppendLine(new string('-', 55));
                foreach (string e in _allErrors)
                    sb.AppendLine("  - " + e);
                sb.AppendLine();
            }

            if (_allErrors.Count == 0)
            {
                sb.AppendLine("No errors.");
            }

            return sb.ToString();
        }

        // ════════════════════════════════════════════════════════════
        //  STEP A – Standardize properties + group by composite key
        // ════════════════════════════════════════════════════════════

        private void StepA_StandardizeAndGroup(Document doc,
            out List<GroupKey> groupKeys,
            out Dictionary<string, List<GroupItem>> groups,
            out Dictionary<string, GroupItem> keepers)
        {
            // ── Collect linear DimensionTypes ──────────────────────
            var dimTypes = new List<DimensionType>();

            foreach (Element elem in new FilteredElementCollector(doc)
                .OfClass(typeof(DimensionType)))
            {
                var dt = elem as DimensionType;
                if (dt == null) continue;

                // Only linear dimensions
                try
                {
                    if (dt.StyleType != DimensionStyleType.Linear)
                        continue;
                }
                catch { continue; }

                // Must have TEXT_FONT to be a styled dimension
                Parameter pFont = dt.get_Parameter(
                    BuiltInParameter.TEXT_FONT);
                if (pFont == null) continue;

                dimTypes.Add(dt);
            }

            // ── Count instances per type ───────────────────────────
            var instanceCount = new Dictionary<int, int>();

            foreach (Dimension inst in new FilteredElementCollector(doc)
                .OfClass(typeof(Dimension))
                .WhereElementIsNotElementType())
            {
                int tid = inst.GetTypeId().IntegerValue;
                if (!instanceCount.ContainsKey(tid))
                    instanceCount[tid] = 0;
                instanceCount[tid]++;
            }

            // ── Collect one sample instance per type ───────────────
            var samplePerType = new Dictionary<int, Dimension>();

            foreach (Dimension inst in new FilteredElementCollector(doc)
                .OfClass(typeof(Dimension))
                .WhereElementIsNotElementType())
            {
                int tid = inst.GetTypeId().IntegerValue;
                if (!samplePerType.ContainsKey(tid))
                    samplePerType[tid] = inst;
            }

            // ── Standardize properties (Transaction) ───────────────
            using (var tx = new Transaction(doc,
                "HMV Dim Audit – Standardize Properties"))
            {
                tx.Start();

                foreach (DimensionType dt in dimTypes)
                {
                    string name = dt.Name;
                    int eid = dt.Id.IntegerValue;
                    var props = new List<string>();

                    TrySet_Font(dt, props);
                    TrySet_Width(dt, props);
                    TrySet_Bold(dt, props);
                    TrySet_Background(dt, props);
                    TrySet_Size(dt, props);

                    if (props.Count > 0)
                    {
                        _typesStandardized++;
                        _propertyChanges.Add(string.Format(
                            "{0} (ID: {1})\n   {2}",
                            name, eid,
                            string.Join("\n   ", props)));
                    }
                }

                tx.Commit();
            }

            // ── Extract unit + decimals and build groups ───────────
            groups = new Dictionary<string, List<GroupItem>>();
            var keyMap = new Dictionary<string, GroupKey>();

            foreach (DimensionType dt in dimTypes)
            {
                string unitStr;
                int decimals;
                ExtractFormatInfo(dt, doc, samplePerType,
                    out unitStr, out decimals);

                double sizeMm = GetSizeMm(dt);
                double snapped = NearestSize(sizeMm);

                var gk = new GroupKey
                {
                    FamilyName = dt.FamilyName,
                    UnitStr = unitStr,
                    SizeMm = snapped,
                    Decimals = decimals
                };
                string key = gk.ToKey();

                if (!groups.ContainsKey(key))
                {
                    groups[key] = new List<GroupItem>();
                    keyMap[key] = gk;
                }

                int count = 0;
                instanceCount.TryGetValue(
                    dt.Id.IntegerValue, out count);

                groups[key].Add(new GroupItem
                {
                    Id = dt.Id.IntegerValue,
                    Name = dt.Name,
                    Count = count
                });
            }

            // ── Build ordered key list ─────────────────────────────
            groupKeys = keyMap.Values.ToList();

            // ── Find keepers per group ─────────────────────────────
            keepers = new Dictionary<string, GroupItem>();

            foreach (var kvp in groups)
            {
                string key = kvp.Key;
                var group = kvp.Value;
                var gk = keyMap[key];
                string stdName = BuildStandardName(gk);

                if (group.Count == 0)
                {
                    keepers[key] = null;
                    continue;
                }

                // Prefer type already named correctly
                GroupItem keeper = group.FirstOrDefault(
                    g => g.Name == stdName);

                // Otherwise pick the one with most instances
                if (keeper == null)
                    keeper = group.OrderByDescending(
                        g => g.Count).First();

                keepers[key] = keeper;
            }
        }

        // ════════════════════════════════════════════════════════════
        //  STEP B – Reassign instances to keepers
        // ════════════════════════════════════════════════════════════

        private void StepB_ReassignInstances(Document doc,
            List<GroupKey> groupKeys,
            Dictionary<string, List<GroupItem>> groups,
            Dictionary<string, GroupItem> keepers,
            out List<int> typesToDelete)
        {
            typesToDelete = new List<int>();

            // Map type → instance IDs (store IDs, not live refs)
            var typeToInstanceIds =
                new Dictionary<int, List<int>>();

            foreach (Dimension inst in
                new FilteredElementCollector(doc)
                    .OfClass(typeof(Dimension))
                    .WhereElementIsNotElementType())
            {
                int tid = inst.GetTypeId().IntegerValue;
                if (!typeToInstanceIds.ContainsKey(tid))
                    typeToInstanceIds[tid] = new List<int>();
                typeToInstanceIds[tid].Add(
                    inst.Id.IntegerValue);
            }

            using (var tx = new Transaction(doc,
                "HMV Dim Audit – Reassign Instances"))
            {
                tx.Start();

                foreach (GroupKey gk in groupKeys)
                {
                    string key = gk.ToKey();
                    var group = groups[key];
                    GroupItem keeperItem = keepers[key];
                    string stdName = BuildStandardName(gk);

                    if (keeperItem == null || group.Count == 0)
                        continue;

                    // Find keeper element (re-fetch fresh)
                    Element keeperElem = doc.GetElement(
                        new ElementId(keeperItem.Id));

                    if (keeperElem == null)
                    {
                        _allErrors.Add(
                            $"{stdName}: Keeper ID " +
                            $"{keeperItem.Id} not found");
                        continue;
                    }

                    // Check if standard name already exists
                    Element stdElem = null;
                    foreach (GroupItem gi in group)
                    {
                        Element e = doc.GetElement(
                            new ElementId(gi.Id));
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
                                var dtKeeper =
                                    keeperElem as DimensionType;
                                if (dtKeeper != null)
                                {
                                    var dup = dtKeeper.Duplicate(
                                        stdName);
                                    finalKeeper = dup;
                                }
                                else
                                {
                                    _allErrors.Add(
                                        $"{stdName}: Could not " +
                                        "rename or duplicate");
                                    continue;
                                }
                            }
                            catch (Exception ex)
                            {
                                _allErrors.Add(
                                    $"{stdName}: {ex.Message}");
                                continue;
                            }
                        }
                    }

                    ElementId finalKeeperId = finalKeeper.Id;
                    int reassigned = 0;
                    int merged = 0;

                    foreach (GroupItem gi in group)
                    {
                        if (gi.Id == finalKeeperId.IntegerValue)
                            continue;

                        List<int> instIds;
                        if (!typeToInstanceIds.TryGetValue(
                            gi.Id, out instIds))
                            instIds = new List<int>();

                        foreach (int instIdInt in instIds)
                        {
                            // Re-fetch fresh from doc each time
                            Dimension inst = doc.GetElement(
                                new ElementId(instIdInt))
                                as Dimension;

                            if (inst == null)
                                continue;

                            // Skip dimensions inside groups
                            if (inst.GroupId !=
                                ElementId.InvalidElementId)
                            {
                                _instancesSkippedGroup++;
                                continue;
                            }

                            try
                            {
                                inst.ChangeTypeId(finalKeeperId);
                                reassigned++;
                            }
                            catch (Exception ex)
                            {
                                _allErrors.Add(
                                    $"Instance {instIdInt}: " +
                                    $"{ex.Message}");
                            }
                        }

                        merged++;
                        typesToDelete.Add(gi.Id);
                    }

                    _instancesReassigned += reassigned;

                    if (merged > 0 || reassigned > 0)
                    {
                        _mergeDetails.Add(
                            $"  [{gk.FamilyName}] {gk.UnitStr} | " +
                            $"{gk.SizeMm}mm | " +
                            $"{gk.Decimals} dec " +
                            $"-> \"{stdName}\": " +
                            $"{merged} types merged, " +
                            $"{reassigned} instances reassigned");
                    }
                }

                tx.Commit();
            }
        }

        // ════════════════════════════════════════════════════════════
        //  STEP C – Delete empty types
        // ════════════════════════════════════════════════════════════

        private void StepC_DeleteEmptyTypes(Document doc,
            List<int> typesToDelete)
        {
            if (typesToDelete == null || typesToDelete.Count == 0)
                return;

            // Recount instances
            var typeCount = new Dictionary<int, int>();

            foreach (Dimension inst in
                new FilteredElementCollector(doc)
                    .OfClass(typeof(Dimension))
                    .WhereElementIsNotElementType())
            {
                int tid = inst.GetTypeId().IntegerValue;
                if (!typeCount.ContainsKey(tid))
                    typeCount[tid] = 0;
                typeCount[tid]++;
            }

            using (var tx = new Transaction(doc,
                "HMV Dim Audit – Delete Empty Types"))
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
                            $"  SKIPPED: {name} (ID: {eid}) " +
                            $"- still has {count} instances");
                        continue;
                    }

                    try
                    {
                        doc.Delete(elem.Id);
                        _typesDeleted++;
                        _deleteDetails.Add(
                            $"  DELETED: {name} (ID: {eid})");
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

        // ════════════════════════════════════════════════════════════
        //  Unit & Decimal extraction via ValueString
        // ════════════════════════════════════════════════════════════

        /// <summary>
        /// Reads unit (m/mm) and decimal count (2 or 3) from a
        /// placed Dimension's ValueString. Falls back to project
        /// units when the type has no instances.
        /// </summary>
        private static void ExtractFormatInfo(
            DimensionType dt, Document doc,
            Dictionary<int, Dimension> samplePerType,
            out string unitStr, out int decimals)
        {
            unitStr = "m";
            decimals = 2;

            Dimension sample;
            if (samplePerType.TryGetValue(
                dt.Id.IntegerValue, out sample))
            {
                // Read ValueString from segment or dimension
                string vs = null;

                try
                {
                    if (sample.NumberOfSegments > 0)
                    {
                        foreach (DimensionSegment seg
                            in sample.Segments)
                        {
                            vs = seg.ValueString;
                            if (!string.IsNullOrEmpty(vs))
                                break;
                        }
                    }
                    else
                    {
                        vs = sample.ValueString;
                    }
                }
                catch { /* skip */ }

                if (!string.IsNullOrEmpty(vs))
                {
                    vs = vs.Trim();

                    // ── Unit detection ──────────────────────
                    if (vs.EndsWith("mm"))
                        unitStr = "mm";
                    else if (vs.EndsWith("m"))
                        unitStr = "m";

                    // ── Decimal detection ───────────────────
                    // Strip unit suffix to isolate numeric part
                    string numeric = vs;
                    if (numeric.EndsWith("mm"))
                        numeric = numeric.Substring(
                            0, numeric.Length - 2).Trim();
                    else if (numeric.EndsWith("m"))
                        numeric = numeric.Substring(
                            0, numeric.Length - 1).Trim();

                    // Handle both '.' and ',' as decimal sep
                    int dotIdx = numeric.LastIndexOf('.');
                    if (dotIdx < 0)
                        dotIdx = numeric.LastIndexOf(',');

                    if (dotIdx >= 0)
                    {
                        int decCount =
                            numeric.Length - dotIdx - 1;
                        decimals = (decCount >= 3) ? 3 : 2;
                    }
                    else
                    {
                        decimals = 2;
                    }
                    return;
                }
            }

            // ── Fallback: project-level units ──────────────────
            try
            {
                FormatOptions fo = doc.GetUnits()
                    .GetFormatOptions(SpecTypeId.Length);
                ForgeTypeId uid = fo.GetUnitTypeId();

                if (uid == UnitTypeId.Millimeters)
                    unitStr = "mm";
                else
                    unitStr = "m";

                double acc = fo.Accuracy;
                decimals = (acc > 0 && acc <= 0.005) ? 3 : 2;
            }
            catch
            {
                unitStr = "m";
                decimals = 2;
            }
        }

        // ════════════════════════════════════════════════════════════
        //  Helpers
        // ════════════════════════════════════════════════════════════

        private static string BuildStandardName(GroupKey gk)
        {
            string xStr = new string('X', gk.Decimals);
            return string.Format(STD_NAME_TEMPLATE,
                gk.UnitStr,
                gk.SizeMm.ToString("0.0"),
                xStr);
        }

        private static double GetSizeMm(DimensionType dt)
        {
            try
            {
                Parameter p = dt.get_Parameter(
                    BuiltInParameter.TEXT_SIZE);
                if (p != null)
                    return Math.Round(p.AsDouble() * 304.8, 4);
            }
            catch { /* skip */ }
            return 2.5; // safe default
        }

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

        // ── Parameter setters ──────────────────────────────────────

        private static void TrySet_Font(Element t,
            List<string> props)
        {
            try
            {
                Parameter p = t.get_Parameter(
                    BuiltInParameter.TEXT_FONT);
                if (p != null && !p.IsReadOnly &&
                    p.AsString() != STD_FONT)
                {
                    string old = p.AsString();
                    p.Set(STD_FONT);
                    props.Add($"Font: {old} -> {STD_FONT}");
                }
            }
            catch { /* skip */ }
        }

        private static void TrySet_Width(Element t,
            List<string> props)
        {
            try
            {
                Parameter p = t.get_Parameter(
                    BuiltInParameter.TEXT_WIDTH_SCALE);
                if (p != null && !p.IsReadOnly &&
                    Math.Round(p.AsDouble(), 4) != STD_WIDTH)
                {
                    double old = Math.Round(p.AsDouble(), 4);
                    p.Set(STD_WIDTH);
                    props.Add($"Width: {old} -> {STD_WIDTH}");
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

        // ── Internal models ─────────────────────────────────────────

        private class GroupKey
        {
            public string FamilyName { get; set; }
            public string UnitStr { get; set; }
            public double SizeMm { get; set; }
            public int Decimals { get; set; }

            public string ToKey()
            {
                return $"{FamilyName}|{UnitStr}|{SizeMm}|{Decimals}";
            }
        }

        private class GroupItem
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public int Count { get; set; }
        }
    }
}