using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class MultiParamEditorCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            ICollection<ElementId> selIds = uidoc.Selection.GetElementIds();
            var instances = new List<FamilyInstance>();
            foreach (ElementId id in selIds)
            {
                Element el = doc.GetElement(id);
                if (el is FamilyInstance fi)
                    instances.Add(fi);
            }

            if (instances.Count == 0)
            {
                TaskDialog.Show("HMV Tools",
                    "Select one or more family instances first.");
                return Result.Cancelled;
            }

            // Group instances by family name
            var familyGroups = new Dictionary<string, List<FamilyInstance>>();
            foreach (FamilyInstance fi in instances)
            {
                string famName = fi.Symbol.Family.Name;
                if (!familyGroups.ContainsKey(famName))
                    familyGroups[famName] = new List<FamilyInstance>();
                familyGroups[famName].Add(fi);
            }

            // Build parameter info per family
            var familyParams =
                new Dictionary<string, List<ParamInfo>>();
            foreach (var kvp in familyGroups)
            {
                var pList = CollectInstanceParams(
                    kvp.Value, doc);
                familyParams[kvp.Key] = pList;
            }

            bool commonMode = false;
            List<string> familyNames = familyGroups.Keys.ToList();

            if (familyNames.Count > 1)
            {
                var modeWin = new MultiParamModeWindow(familyNames);
                if (modeWin.ShowDialog() != true)
                    return Result.Cancelled;
                commonMode = modeWin.IsCommonMode;
            }

            Dictionary<string, List<ParamInfo>> editorParams;
            if (commonMode)
            {
                var common = GetCommonParams(familyParams);
                editorParams = new Dictionary<string, List<ParamInfo>>
                {
                    { "Common", common }
                };
            }
            else
            {
                editorParams = familyParams;
            }

            var editorWin = new MultiParamEditorWindow(
                editorParams, familyGroups, commonMode);
            if (editorWin.ShowDialog() != true)
                return Result.Cancelled;

            // Apply changes
            var changes = editorWin.ResultChanges;
            if (changes == null || changes.Count == 0)
                return Result.Cancelled;

            int changed = 0;
            int errors = 0;

            using (Transaction tx = new Transaction(doc,
                "HMV – Multi-Parameter Edit"))
            {
                tx.Start();

                foreach (var entry in changes)
                {
                    string famKey = entry.Key;
                    List<FamilyInstance> targets;

                    if (commonMode)
                    {
                        targets = instances;
                    }
                    else
                    {
                        if (!familyGroups.ContainsKey(famKey))
                            continue;
                        targets = familyGroups[famKey];
                    }

                    foreach (var change in entry.Value)
                    {
                        foreach (FamilyInstance fi in targets)
                        {
                            Parameter p =
                                fi.LookupParameter(change.ParamName);
                            if (p == null || p.IsReadOnly)
                                continue;

                            try
                            {
                                if (p.StorageType == StorageType.Double)
                                {
                                    double feet = ConvertToInternal(
                                        change.NewValue, p, doc);
                                    p.Set(feet);
                                    changed++;
                                }
                                else if (p.StorageType == StorageType.Integer)
                                {
                                    if (int.TryParse(change.NewValue,
                                        out int iv))
                                    {
                                        p.Set(iv);
                                        changed++;
                                    }
                                    else errors++;
                                }
                                else if (p.StorageType == StorageType.String)
                                {
                                    p.Set(change.NewValue);
                                    changed++;
                                }
                            }
                            catch
                            {
                                errors++;
                            }
                        }
                    }
                }

                tx.Commit();
            }

            TaskDialog.Show("HMV Tools – Multi-Parameter Editor",
                $"Parameters set: {changed}\n"
              + $"Errors: {errors}");

            return Result.Succeeded;
        }

        private List<ParamInfo> CollectInstanceParams(
            List<FamilyInstance> instances, Document doc)
        {
            var result = new List<ParamInfo>();
            var seen = new HashSet<string>();
            FamilyInstance primary = instances[0];

            foreach (Parameter p in primary.Parameters)
            {
                if (p.IsReadOnly) continue;
                if (p.Definition == null) continue;
                if (!p.IsShared
                    && p.Definition.ParameterGroup
                        == BuiltInParameterGroup.INVALID)
                    continue;

                string name = p.Definition.Name;
                if (seen.Contains(name)) continue;
                if (p.StorageType == StorageType.ElementId)
                    continue;
                if (p.StorageType == StorageType.None)
                    continue;

                seen.Add(name);

                // Collect current values from all instances
                var values = new List<string>();
                foreach (FamilyInstance fi in instances)
                {
                    Parameter fp = fi.LookupParameter(name);
                    if (fp != null)
                        values.Add(FormatValue(fp, doc));
                    else
                        values.Add("");
                }

                bool varies = values.Distinct().Count() > 1;
                string displayUnit = GetDisplayUnit(p, doc);

                result.Add(new ParamInfo
                {
                    Name = name,
                    StorageType = p.StorageType,
                    CurrentValue = varies
                        ? string.Join(", ", values.Distinct())
                        : values[0],
                    Varies = varies,
                    DisplayUnit = displayUnit
                });
            }

            result.Sort((a, b) =>
                string.Compare(a.Name, b.Name,
                    StringComparison.OrdinalIgnoreCase));
            return result;
        }

        private string FormatValue(Parameter p, Document doc)
        {
            if (p.StorageType == StorageType.Double)
            {
                double internalVal = p.AsDouble();
                try
                {
                    ForgeTypeId specId =
                        p.Definition.GetDataType();
                    ForgeTypeId unitId = doc.GetUnits()
                        .GetFormatOptions(specId)
                        .GetUnitTypeId();
                    double converted =
                        UnitUtils.ConvertFromInternalUnits(
                            internalVal, unitId);
                    return converted.ToString("F2");
                }
                catch
                {
                    return internalVal.ToString("F6");
                }
            }
            if (p.StorageType == StorageType.Integer)
                return p.AsInteger().ToString();
            if (p.StorageType == StorageType.String)
                return p.AsString() ?? "";
            return "";
        }

        private string GetDisplayUnit(Parameter p, Document doc)
        {
            if (p.StorageType != StorageType.Double)
                return "";
            try
            {
                ForgeTypeId specId =
                    p.Definition.GetDataType();
                ForgeTypeId unitId = doc.GetUnits()
                    .GetFormatOptions(specId)
                    .GetUnitTypeId();
                if (unitId == UnitTypeId.Millimeters)
                    return "mm";
                if (unitId == UnitTypeId.Meters)
                    return "m";
                if (unitId == UnitTypeId.Centimeters)
                    return "cm";
                if (unitId == UnitTypeId.Inches
                    || unitId == UnitTypeId.FractionalInches)
                    return "in";
                if (unitId == UnitTypeId.Feet
                    || unitId == UnitTypeId.FeetFractionalInches)
                    return "ft";
                if (unitId == UnitTypeId.Degrees)
                    return "°";
                return "";
            }
            catch
            {
                return "";
            }
        }

        private double ConvertToInternal(
            string text, Parameter p, Document doc)
        {
            double val = double.Parse(text);
            try
            {
                ForgeTypeId specId =
                    p.Definition.GetDataType();
                ForgeTypeId unitId = doc.GetUnits()
                    .GetFormatOptions(specId)
                    .GetUnitTypeId();
                return UnitUtils.ConvertToInternalUnits(
                    val, unitId);
            }
            catch
            {
                return val;
            }
        }

        private List<ParamInfo> GetCommonParams(
            Dictionary<string, List<ParamInfo>> familyParams)
        {
            var lists = familyParams.Values.ToList();
            if (lists.Count == 0)
                return new List<ParamInfo>();

            var first = lists[0];
            var common = new List<ParamInfo>();

            foreach (ParamInfo pi in first)
            {
                bool inAll = true;
                for (int i = 1; i < lists.Count; i++)
                {
                    if (!lists[i].Any(p => p.Name == pi.Name
                        && p.StorageType == pi.StorageType))
                    {
                        inAll = false;
                        break;
                    }
                }
                if (inAll)
                    common.Add(pi);
            }

            return common;
        }
    }

    public class ParamInfo
    {
        public string Name;
        public StorageType StorageType;
        public string CurrentValue;
        public bool Varies;
        public string DisplayUnit;
    }

    public class ParamChange
    {
        public string ParamName;
        public string NewValue;
    }
}
