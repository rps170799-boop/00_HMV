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
    public class SpotElevationCommand : IExternalCommand
    {
        private const double MmToFeet = 1.0 / 304.8;
        private const string NtceTypeName = "HMV - NTCE";
        private const string NapTypeName = "HMV - NAP";

        private class FoundationData
        {
            public XYZ HostCenter;
            public XYZ BBoxMin;
            public XYZ BBoxMax;
            public string Name;
            public ElementId ElementId;
            public ElementId LinkInstanceId;
        }

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View view = doc.ActiveView;

            if (!(view is ViewPlan))
            {
                TaskDialog.Show("HMV Tools",
                    "This command only works in plan views.");
                return Result.Cancelled;
            }

            // ── Collect links ───────────────────────────────────
            var linkInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(li => li.GetLinkDocument() != null)
                .ToList();

            if (linkInstances.Count == 0)
            {
                TaskDialog.Show("HMV Tools",
                    "No loaded Revit links found.");
                return Result.Cancelled;
            }

            var linkInfos = new List<LinkInfo>();
            for (int i = 0; i < linkInstances.Count; i++)
                linkInfos.Add(new LinkInfo
                { Name = linkInstances[i].Name, Index = i });

            // ── Settings window ─────────────────────────────────
            var win = new SpotElevationWindow(linkInfos);
            if (win.ShowDialog() != true || win.Settings == null)
                return Result.Cancelled;

            SpotElevationSettings settings = win.Settings;
            double leaderOffset = settings.LeaderOffsetMm * MmToFeet;
            bool bothAxis = settings.BothAxis;
            bool hmvStandard = settings.UseHmvStandard;

            RevitLinkInstance floorLink = linkInstances[settings.FloorLinkIndex];
            Document floorDoc = floorLink.GetLinkDocument();

            // ── 3D view ─────────────────────────────────────────
            View3D view3d = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .FirstOrDefault(v => !v.IsTemplate && !v.IsLocked);

            if (view3d == null)
            {
                TaskDialog.Show("HMV Tools",
                    "A 3D view is required. Create {3D} and retry.");
                return Result.Cancelled;
            }

            // ── Pick elements ───────────────────────────────────
            IList<Reference> pickedRefs;
            bool fromLink = settings.FoundationSourceIndex >= 0;

            try
            {
                pickedRefs = fromLink
                    ? uidoc.Selection.PickObjects(ObjectType.LinkedElement,
                        "Select elements in link. Finish when done.")
                    : uidoc.Selection.PickObjects(ObjectType.Element,
                        "Select elements. Finish when done.");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            { return Result.Cancelled; }

            if (pickedRefs == null || pickedRefs.Count == 0)
                return Result.Cancelled;

            // ── Resolve foundation data ─────────────────────────
            var foundations = new List<FoundationData>();

            foreach (Reference r in pickedRefs)
            {
                Element elem;
                XYZ center, bbMinH, bbMaxH;
                ElementId elemId;
                ElementId linkInstId = ElementId.InvalidElementId;

                if (fromLink)
                {
                    var srcLink = doc.GetElement(r) as RevitLinkInstance;
                    if (srcLink == null) continue;
                    Document srcDoc = srcLink.GetLinkDocument();
                    if (srcDoc == null) continue;
                    elem = srcDoc.GetElement(r.LinkedElementId);
                    if (elem == null) continue;

                    Transform t = srcLink.GetTotalTransform();
                    BoundingBoxXYZ bb = elem.get_BoundingBox(null);
                    if (bb == null) continue;

                    center = t.OfPoint((bb.Min + bb.Max) / 2.0);
                    bbMinH = t.OfPoint(bb.Min);
                    bbMaxH = t.OfPoint(bb.Max);
                    elemId = r.LinkedElementId;
                    linkInstId = srcLink.Id;
                }
                else
                {
                    elem = doc.GetElement(r);
                    if (elem == null) continue;
                    BoundingBoxXYZ bb = elem.get_BoundingBox(null);
                    if (bb == null) continue;

                    center = (bb.Min + bb.Max) / 2.0;
                    bbMinH = bb.Min;
                    bbMaxH = bb.Max;
                    elemId = elem.Id;
                }

                foundations.Add(new FoundationData
                {
                    HostCenter = center,
                    BBoxMin = bbMinH,
                    BBoxMax = bbMaxH,
                    Name = $"{elem.Category?.Name ?? "Element"}: "
                                   + $"{elem.Name} (Id {elem.Id.IntegerValue})",
                    ElementId = elemId,
                    LinkInstanceId = linkInstId
                });
            }

            // ── Intersectors ────────────────────────────────────
            var floorRI = new ReferenceIntersector(
                new ElementCategoryFilter(BuiltInCategory.OST_Floors),
                FindReferenceTarget.Face, view3d);
            floorRI.FindReferencesInRevitLinks = true;

            ReferenceIntersector elemRI = null;
            if (hmvStandard)
            {
                elemRI = new ReferenceIntersector(
                    new ElementIsElementTypeFilter(true),
                    FindReferenceTarget.Face, view3d);
                elemRI.FindReferencesInRevitLinks = true;
            }

            // ── Process ─────────────────────────────────────────
            int placedNap = 0, placedNtce = 0;
            var skipped = new List<string>();
            var debugInfo = new List<string>();

            int viewScale = view.Scale;
            double gapFeet = 2.5 * 3.5 * viewScale * MmToFeet;

            using (Transaction tx = new Transaction(doc,
                "HMV – Spot Elevations on Floors"))
            {
                tx.Start();

                // ── Find or create types ────────────────────────
                ElementId ntceTypeId = ElementId.InvalidElementId;
                ElementId napTypeId = ElementId.InvalidElementId;

                if (hmvStandard)
                {
                    // Find a base type to duplicate from
                    ElementId baseTypeId = FindBaseSpotTypeId(doc);

                    if (baseTypeId == ElementId.InvalidElementId)
                    {
                        debugInfo.Add(
                            "CRITICAL: No base SpotDimensionType found. "
                          + "Types will not be assigned.");
                    }
                    else
                    {
                        ntceTypeId = FindOrCreateHmvType(
                            doc, baseTypeId, NtceTypeName, "N.T.C.E. ");
                        napTypeId = FindOrCreateHmvType(
                            doc, baseTypeId, NapTypeName, "N.A.P. ");

                        if (ntceTypeId == ElementId.InvalidElementId)
                            debugInfo.Add("WARNING: Could not create NTCE type");
                        if (napTypeId == ElementId.InvalidElementId)
                            debugInfo.Add("WARNING: Could not create NAP type");
                    }
                }

                foreach (var fd in foundations)
                {
                    XYZ hostCenter = fd.HostCenter;
                    double offX = leaderOffset;
                    double offY = bothAxis ? leaderOffset : 0;

                    // ── NAP: floor top face ─────────────────────
                    XYZ rayOrigin = new XYZ(
                        hostCenter.X, hostCenter.Y, hostCenter.Z + 200);

                    var floorHits = floorRI.Find(
                        rayOrigin, XYZ.BasisZ.Negate());
                    var napHit = FindFirstHitOnLink(
                        floorHits, floorLink.Id);

                    if (napHit == null)
                    {
                        skipped.Add($"{fd.Name} → No floor face hit");
                        continue;
                    }

                    Reference napRef = napHit.GetReference();
                    XYZ napPoint = napRef.GlobalPoint
                        ?? new XYZ(hostCenter.X, hostCenter.Y,
                            rayOrigin.Z - napHit.Proximity);

                    if (hmvStandard)
                    {
                        // ── NTCE: highest face on element ───────
                        var ntceHit = FindHighestFaceOnElement(
                            elemRI, fd.BBoxMin, fd.BBoxMax,
                            fd.LinkInstanceId, fd.ElementId);

                        if (ntceHit != null)
                        {
                            Reference ntceRef = ntceHit.GetReference();

                            // Compute NTCE point from ray origin
                            // (ray origin is at hostCenter.Z + 200
                            //  for element rays it varies per sample)
                            XYZ ntcePoint = ntceRef.GlobalPoint;
                            if (ntcePoint == null)
                            {
                                // Approximate from bounding box top
                                double topZ = Math.Max(
                                    fd.BBoxMax.Z, fd.BBoxMin.Z);
                                ntcePoint = new XYZ(
                                    hostCenter.X, hostCenter.Y, topZ);
                            }

                            // NTCE: leader ON, upper position
                            XYZ ntceBend = new XYZ(
                                ntcePoint.X + offX,
                                ntcePoint.Y + offY,
                                ntcePoint.Z);
                            XYZ ntceEnd = new XYZ(
                                ntceBend.X + offX * 0.5,
                                ntceBend.Y,
                                ntceBend.Z);

                            try
                            {
                                SpotDimension spotNtce =
                                    doc.Create.NewSpotElevation(
                                        view, ntceRef,
                                        ntcePoint, ntceBend, ntceEnd,
                                        ntcePoint, true);

                                if (spotNtce != null)
                                {
                                    if (ntceTypeId != ElementId.InvalidElementId)
                                    {
                                        try { spotNtce.ChangeTypeId(ntceTypeId); }
                                        catch (Exception ex)
                                        { debugInfo.Add($"NTCE type: {ex.Message}"); }
                                    }
                                    placedNtce++;
                                }
                            }
                            catch (Exception ex)
                            {
                                skipped.Add(
                                    $"{fd.Name} → NTCE error: {ex.Message}");
                            }
                        }
                        else
                        {
                            skipped.Add(
                                $"{fd.Name} → NTCE: no top face found "
                              + $"(multi-ray on bbox)");
                        }

                        // ── NAP: leader ON, below NTCE ──────────
                        XYZ napBend = new XYZ(
                            napPoint.X + offX,
                            napPoint.Y + offY - gapFeet,
                            napPoint.Z);
                        XYZ napEnd = new XYZ(
                            napBend.X + offX * 0.5,
                            napBend.Y,
                            napBend.Z);

                        try
                        {
                            SpotDimension spotNap =
                                doc.Create.NewSpotElevation(
                                    view, napRef,
                                    napPoint, napBend, napEnd,
                                    napPoint, true);

                            if (spotNap != null)
                            {
                                if (napTypeId != ElementId.InvalidElementId)
                                {
                                    try { spotNap.ChangeTypeId(napTypeId); }
                                    catch (Exception ex)
                                    { debugInfo.Add($"NAP type: {ex.Message}"); }
                                }
                                placedNap++;
                            }
                        }
                        catch (Exception ex)
                        {
                            skipped.Add(
                                $"{fd.Name} → NAP error: {ex.Message}");
                        }
                    }
                    else
                    {
                        // ── Simple mode ─────────────────────────
                        XYZ bend = new XYZ(
                            napPoint.X + offX,
                            napPoint.Y + offY,
                            napPoint.Z);
                        XYZ end = new XYZ(
                            bend.X + offX * 0.5,
                            bend.Y,
                            bend.Z);

                        try
                        {
                            SpotDimension spot =
                                doc.Create.NewSpotElevation(
                                    view, napRef,
                                    napPoint, bend, end,
                                    napPoint, true);

                            if (spot != null) placedNap++;
                            else skipped.Add(
                                $"{fd.Name} → returned null");
                        }
                        catch (Exception ex)
                        {
                            skipped.Add(
                                $"{fd.Name} → {ex.Message}");
                        }
                    }
                }

                tx.Commit();
            }

            // ── Report ──────────────────────────────────────────
            string report = "SPOT ELEVATION SUMMARY\n"
                + "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
                + $"Elements picked: {foundations.Count}\n";

            if (hmvStandard)
                report += $"NTCE placed: {placedNtce}\n"
                        + $"NAP placed: {placedNap}\n";
            else
                report += $"Spot elevations placed: {placedNap}\n";

            report += $"Skipped: {skipped.Count}\n";

            if (skipped.Count > 0)
            {
                report += "\n── SKIPPED ──\n";
                foreach (string s in skipped)
                    report += $"  • {s}\n";
            }

            if (debugInfo.Count > 0)
            {
                report += "\n── DEBUG INFO ──\n";
                foreach (string s in debugInfo)
                    report += $"  ▸ {s}\n";
            }

            var td = new TaskDialog("HMV Tools – Spot Elevation");
            td.MainContent = report;
            td.Show();

            return Result.Succeeded;
        }

        // ════════════════════════════════════════════════════════
        // Multi-ray: find HIGHEST face on element
        // 5×5 grid + extra corner insets = ~37 rays
        // ════════════════════════════════════════════════════════

        private ReferenceWithContext FindHighestFaceOnElement(
            ReferenceIntersector intersector,
            XYZ bbMin, XYZ bbMax,
            ElementId linkInstanceId,
            ElementId elementId)
        {
            double xMin = Math.Min(bbMin.X, bbMax.X);
            double xMax = Math.Max(bbMin.X, bbMax.X);
            double yMin = Math.Min(bbMin.Y, bbMax.Y);
            double yMax = Math.Max(bbMin.Y, bbMax.Y);
            double zTop = Math.Max(bbMin.Z, bbMax.Z);

            double dx = xMax - xMin;
            double dy = yMax - yMin;

            // Build sample points: 5×5 grid
            var sampleXY = new List<XYZ>();
            for (int ix = 0; ix < 5; ix++)
            {
                double px = xMin + dx * (0.05 + 0.9 * ix / 4.0);
                for (int iy = 0; iy < 5; iy++)
                {
                    double py = yMin + dy * (0.05 + 0.9 * iy / 4.0);
                    sampleXY.Add(new XYZ(px, py, 0));
                }
            }

            // Extra: corners at multiple insets (0.02, 0.10, 0.15)
            double[] insets = { 0.02, 0.10, 0.15 };
            foreach (double f in insets)
            {
                double ix = dx * f;
                double iy = dy * f;
                sampleXY.Add(new XYZ(xMin + ix, yMin + iy, 0));
                sampleXY.Add(new XYZ(xMax - ix, yMin + iy, 0));
                sampleXY.Add(new XYZ(xMin + ix, yMax - iy, 0));
                sampleXY.Add(new XYZ(xMax - ix, yMax - iy, 0));
            }

            ReferenceWithContext bestHit = null;
            double bestProximity = double.MaxValue;
            XYZ direction = XYZ.BasisZ.Negate();

            foreach (XYZ xy in sampleXY)
            {
                XYZ origin = new XYZ(xy.X, xy.Y, zTop + 200);

                IList<ReferenceWithContext> hits =
                    intersector.Find(origin, direction);
                if (hits == null) continue;

                foreach (var rwc in hits)
                {
                    Reference r = rwc.GetReference();
                    bool match;

                    if (linkInstanceId == ElementId.InvalidElementId)
                        match = r.ElementId == elementId;
                    else
                        match = r.ElementId == linkInstanceId
                             && r.LinkedElementId == elementId;

                    if (!match) continue;

                    // Smallest proximity = highest face
                    if (rwc.Proximity < bestProximity)
                    {
                        bestProximity = rwc.Proximity;
                        bestHit = rwc;
                    }
                }
            }

            return bestHit;
        }

        // ════════════════════════════════════════════════════════
        // Find first hit on a link
        // ════════════════════════════════════════════════════════

        private ReferenceWithContext FindFirstHitOnLink(
            IList<ReferenceWithContext> hits, ElementId linkId)
        {
            if (hits == null) return null;
            ReferenceWithContext best = null;
            double bestProx = double.MaxValue;
            foreach (var rwc in hits)
            {
                if (rwc.GetReference().ElementId != linkId) continue;
                if (rwc.Proximity < bestProx)
                { bestProx = rwc.Proximity; best = rwc; }
            }
            return best;
        }

        // ════════════════════════════════════════════════════════
        // Find a base SpotDimensionType ID to duplicate from
        // ════════════════════════════════════════════════════════

        /// <summary>
        /// Strategy:
        /// 1) GetDefaultElementTypeId(SpotElevationType)
        /// 2) Find an existing SpotDimension instance, get its TypeId
        /// 3) Iterate all DimensionType elements, find one whose
        ///    FamilyName contains "Spot" or "Elevation"
        /// </summary>
        private ElementId FindBaseSpotTypeId(Document doc)
        {
            // Approach 1: default
            try
            {
                ElementId defId = doc.GetDefaultElementTypeId(
                    ElementTypeGroup.SpotElevationType);
                if (defId != null && defId != ElementId.InvalidElementId)
                    return defId;
            }
            catch { }

            // Approach 2: from existing instance
            try
            {
                var existingSpot = new FilteredElementCollector(doc)
                    .OfClass(typeof(SpotDimension))
                    .WhereElementIsNotElementType()
                    .FirstOrDefault() as SpotDimension;

                if (existingSpot != null)
                {
                    ElementId tid = existingSpot.GetTypeId();
                    if (tid != null && tid != ElementId.InvalidElementId)
                        return tid;
                }
            }
            catch { }

            // Approach 3: search DimensionTypes
            try
            {
                var dimTypes = new FilteredElementCollector(doc)
                    .OfClass(typeof(DimensionType))
                    .WhereElementIsElementType()
                    .ToElements();

                foreach (Element dt in dimTypes)
                {
                    string famName = (dt as DimensionType)?
                        .FamilyName ?? "";
                    if (famName.IndexOf("Spot",
                        StringComparison.OrdinalIgnoreCase) >= 0
                     || famName.IndexOf("Elevation",
                        StringComparison.OrdinalIgnoreCase) >= 0
                     || famName.IndexOf("Cota",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return dt.Id;
                    }
                }
            }
            catch { }

            return ElementId.InvalidElementId;
        }

        // ════════════════════════════════════════════════════════
        // Find or create an HMV SpotDimension type
        // ════════════════════════════════════════════════════════

        private ElementId FindOrCreateHmvType(
            Document doc,
            ElementId baseTypeId,
            string typeName,
            string prefix)
        {
            // Check if already exists (by name across all DimensionTypes)
            try
            {
                // Search all element types for matching name
                var allTypes = new FilteredElementCollector(doc)
                    .WhereElementIsElementType()
                    .ToElements();

                foreach (Element e in allTypes)
                {
                    if (e.Name == typeName)
                        return e.Id;
                }
            }
            catch { }

            // Duplicate from base
            ElementType baseType =
                doc.GetElement(baseTypeId) as ElementType;
            if (baseType == null) return ElementId.InvalidElementId;

            try
            {
                ElementType newType = baseType.Duplicate(typeName);
                if (newType == null) return ElementId.InvalidElementId;

                // Text: Arial 2.5mm, no bold, opaque, width 1.0
                TrySet(newType, BuiltInParameter.TEXT_SIZE,
                    2.5 * MmToFeet);
                TrySet(newType, BuiltInParameter.TEXT_FONT, "Arial");
                TrySet(newType, BuiltInParameter.TEXT_STYLE_BOLD, 0);
                TrySet(newType, BuiltInParameter.TEXT_BACKGROUND, 0);
                TrySet(newType, BuiltInParameter.TEXT_WIDTH_SCALE, 1.0);

                // Prefix on the TYPE
                TrySet(newType,
                    BuiltInParameter.SPOT_ELEV_SINGLE_OR_UPPER_PREFIX,
                    prefix);

                // Text location: Above Leader
                TrySet(newType,
                    BuiltInParameter.SPOT_ELEV_TEXT_LOCATION, 0);

                // Display Elevations: Actual (Selected) = survey point
                TrySet(newType,
                    BuiltInParameter.SPOT_ELEV_DISPLAY_ELEVATIONS, 0);

                return newType.Id;
            }
            catch
            {
                return ElementId.InvalidElementId;
            }
        }

        // ── Safe parameter setters ──────────────────────────────

        private void TrySet(Element e, BuiltInParameter bip, double v)
        {
            try
            {
                Parameter p = e.get_Parameter(bip);
                if (p != null && !p.IsReadOnly) p.Set(v);
            }
            catch { }
        }

        private void TrySet(Element e, BuiltInParameter bip, int v)
        {
            try
            {
                Parameter p = e.get_Parameter(bip);
                if (p != null && !p.IsReadOnly) p.Set(v);
            }
            catch { }
        }

        private void TrySet(Element e, BuiltInParameter bip, string v)
        {
            try
            {
                Parameter p = e.get_Parameter(bip);
                if (p != null && !p.IsReadOnly) p.Set(v);
            }
            catch { }
        }
    }
}