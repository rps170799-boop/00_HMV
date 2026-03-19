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
        private const string NtceTypeName =
            "HMV - Coordenadas Verticales (Globales) - N.T.C.E.";
        private const string NapTypeName =
            "HMV - Coordenadas Verticales (Globales) - N.A.P.";

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

            var win = new SpotElevationWindow(linkInfos);
            if (win.ShowDialog() != true || win.Settings == null)
                return Result.Cancelled;

            SpotElevationSettings settings = win.Settings;
            double leaderOffset = settings.LeaderOffsetMm * MmToFeet;
            bool bothAxis = settings.BothAxis;
            bool hmvStandard = settings.UseHmvStandard;

            RevitLinkInstance floorLink =
                linkInstances[settings.FloorLinkIndex];

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

            // Collect placed spots for type assignment later
            var ntceSpots = new List<SpotDimension>();
            var napSpots = new List<SpotDimension>();

            using (Transaction tx = new Transaction(doc,
                "HMV – Spot Elevations on Floors"))
            {
                tx.Start();

                // ════════════════════════════════════════════════
                // PHASE 1: Place all spot elevations
                // ════════════════════════════════════════════════

                foreach (var fd in foundations)
                {
                    XYZ hostCenter = fd.HostCenter;
                    double offX = leaderOffset;
                    double offY = bothAxis ? leaderOffset : 0;

                    // ── NAP: floor top face ─────────────────────
                    XYZ rayOrigin = new XYZ(
                        hostCenter.X, hostCenter.Y,
                        hostCenter.Z + 200);

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
                            XYZ ntcePoint = ntceRef.GlobalPoint;
                            if (ntcePoint == null)
                            {
                                double topZ = Math.Max(
                                    fd.BBoxMax.Z, fd.BBoxMin.Z);
                                ntcePoint = new XYZ(
                                    hostCenter.X, hostCenter.Y, topZ);
                            }

                            // NTCE: create WITH leader at same position
                            // as NAP, then disable leader. Text stays at
                            // its positioned location above the shoulder.
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
                                // Create with leader ON (white = invisible)
                                // Same bend/end as NAP for alignment
                                SpotDimension spotNtce =
                                    doc.Create.NewSpotElevation(
                                        view, ntceRef,
                                        ntcePoint, ntceBend, ntceEnd,
                                        ntcePoint, true);

                                if (spotNtce != null)
                                {
                                    ntceSpots.Add(spotNtce);
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
                                $"{fd.Name} → NTCE: no top face found");
                        }

                        // ── NAP: leader ON, text Below Leader ────────
                        // Same bend offset → shoulder lines overlap,
                        // NTCE text Above + NAP text Below = stacked
                        XYZ napBend = new XYZ(
                            napPoint.X + offX,
                            napPoint.Y + offY,
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
                                XYZ nloc = (spotNap.Location as LocationPoint)?.Point;
                                string nlocStr = nloc != null
                                    ? $"({nloc.X:F2},{nloc.Y:F2},{nloc.Z:F2})"
                                    : "null";
                                debugInfo.Add(
                                    $"NAP placed: origin=({napPoint.X:F2},{napPoint.Y:F2},{napPoint.Z:F2})"
                                  + $" bend=({napBend.X:F2},{napBend.Y:F2})"
                                  + $" loc={nlocStr}"
                                  + $" HasLeader={spotNap.HasLeader}");

                                napSpots.Add(spotNap);
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

                // ════════════════════════════════════════════════
                // PHASE 2: Create types from placed spot's type
                // and assign to all placed spots
                // ════════════════════════════════════════════════

                if (hmvStandard)
                {
                    // Get a base type from any placed spot
                    SpotDimension anySpot =
                        ntceSpots.FirstOrDefault()
                        ?? napSpots.FirstOrDefault();

                    if (anySpot != null)
                    {
                        ElementId baseTypeId = anySpot.GetTypeId();
                        ElementType baseType =
                            doc.GetElement(baseTypeId) as ElementType;

                        debugInfo.Add(
                            $"Base type: '{baseType?.Name}' "
                          + $"(Id {baseTypeId.IntegerValue})");

                        if (baseType != null)
                        {
                            // Find or create NTCE type
                            // (text Above Leader, white leader = invisible)
                            ElementId ntceTypeId =
                                FindOrCreateType(
                                    doc, baseType,
                                    NtceTypeName, "N.T.C.E.",
                                    false, true,
                                    debugInfo);

                            // Find or create NAP type
                            // (text Below Leader, black leader = visible)
                            ElementId napTypeId =
                                FindOrCreateType(
                                    doc, baseType,
                                    NapTypeName, "N.A.P.",
                                    true, false,
                                    debugInfo);

                            debugInfo.Add(
                                $"NTCE type Id: {ntceTypeId.IntegerValue}, "
                              + $"NAP type Id: {napTypeId.IntegerValue}");

                            // Assign NTCE type
                            if (ntceTypeId != ElementId.InvalidElementId)
                            {
                                foreach (var s in ntceSpots)
                                {
                                    try { s.ChangeTypeId(ntceTypeId); }
                                    catch (Exception ex)
                                    {
                                        debugInfo.Add(
                                            $"NTCE ChangeTypeId: {ex.Message}");
                                    }
                                }
                            }

                            // Assign NAP type
                            if (napTypeId != ElementId.InvalidElementId)
                            {
                                foreach (var s in napSpots)
                                {
                                    try { s.ChangeTypeId(napTypeId); }
                                    catch (Exception ex)
                                    {
                                        debugInfo.Add(
                                            $"NAP ChangeTypeId: {ex.Message}");
                                    }
                                }
                            }
                        }
                        else
                        {
                            debugInfo.Add(
                                "CRITICAL: baseType is null");
                        }
                    }
                    else
                    {
                        debugInfo.Add(
                            "No spots placed, skipping type creation");
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
                report += "\n── DEBUG ──\n";
                foreach (string s in debugInfo)
                    report += $"  ▸ {s}\n";
            }

            var td = new TaskDialog("HMV Tools – Spot Elevation");
            td.MainContent = report;
            td.Show();

            return Result.Succeeded;
        }

        // ════════════════════════════════════════════════════════
        // Find or create HMV type by duplicating from base type
        // (same pattern as USDTuberiaConfigCommand.FindOrCreateType)
        // ════════════════════════════════════════════════════════

        private ElementId FindOrCreateType(
            Document doc,
            ElementType baseType,
            string typeName,
            string prefix,
            bool textBelowLeader,
            bool whiteLeader,
            List<string> debugInfo)
        {
            // Check if it already exists
            var allTypes = new FilteredElementCollector(doc)
                .WhereElementIsElementType()
                .ToElements();

            foreach (Element e in allTypes)
            {
                if (e.Name == typeName
                    && e.GetType().Name == baseType.GetType().Name)
                    return e.Id;
            }

            // Duplicate from base
            try
            {
                ElementType newType = baseType.Duplicate(typeName);
                if (newType == null)
                    return ElementId.InvalidElementId;

                // Text: Arial 2.5mm, no bold, opaque, width 1.0
                TrySet(newType, BuiltInParameter.TEXT_SIZE,
                    2.5 * MmToFeet);
                TrySet(newType, BuiltInParameter.TEXT_FONT,
                    "Arial");
                TrySet(newType, BuiltInParameter.TEXT_STYLE_BOLD,
                    0);
                TrySet(newType, BuiltInParameter.TEXT_BACKGROUND,
                    0);
                TrySet(newType, BuiltInParameter.TEXT_WIDTH_SCALE,
                    1.0);

                // Text location: 0=Above Leader, 1=Below Leader
                TrySet(newType,
                    BuiltInParameter.SPOT_ELEV_TEXT_LOCATION,
                    textBelowLeader ? 1 : 0);

                // Display Elevations: Actual (Selected)
                TrySet(newType,
                    BuiltInParameter.SPOT_ELEV_DISPLAY_ELEVATIONS,
                    0);

                // ── Set parameters by name (language-independent) ──
                bool foundIndicator = false;
                bool foundBase = false;
                bool foundColor = false;
                var typeParamNames = new List<string>();

                foreach (Parameter p in newType.Parameters)
                {
                    string pName = p.Definition?.Name;
                    if (pName == null) continue;

                    // Collect all RW string/int params for debug
                    if (!p.IsReadOnly)
                        typeParamNames.Add(
                            $"{pName} ({p.StorageType})");

                    if (p.IsReadOnly) continue;

                    // ONLY "Elevation Indicator" — not Bottom/Top
                    if (pName.Equals("Elevation Indicator",
                            StringComparison.OrdinalIgnoreCase)
                        || pName.Equals("Indicador de elevación",
                            StringComparison.OrdinalIgnoreCase)
                        || pName.Equals("Indicador de elevacion",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        if (p.StorageType == StorageType.String)
                        {
                            p.Set(prefix);
                            foundIndicator = true;
                            debugInfo.Add(
                                $"  '{typeName}': set '{pName}' = '{prefix}'");
                        }
                    }

                    // Clear Bottom/Top indicators (set to empty)
                    if (pName.IndexOf("Bottom Indicator",
                            StringComparison.OrdinalIgnoreCase) >= 0
                        || pName.IndexOf("Top Indicator",
                            StringComparison.OrdinalIgnoreCase) >= 0
                        || pName.IndexOf("Indicador inferior",
                            StringComparison.OrdinalIgnoreCase) >= 0
                        || pName.IndexOf("Indicador superior",
                            StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        if (p.StorageType == StorageType.String)
                            p.Set("");
                    }

                    // Leader color → white if whiteLeader
                    if (whiteLeader
                        && pName.Equals("Color",
                                StringComparison.OrdinalIgnoreCase))
                    {
                        if (p.StorageType == StorageType.Integer)
                        {
                            try
                            {
                                p.Set(16777215); // 0xFFFFFF = white
                                foundColor = true;
                            }
                            catch { }
                        }
                    }

                    // Leader Arrowhead → None for white leader
                    if (whiteLeader
                        && (pName.IndexOf("Leader Arrowhead",
                                StringComparison.OrdinalIgnoreCase) >= 0
                            || pName.IndexOf("Punta de flecha",
                                StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        if (p.StorageType == StorageType.ElementId)
                        {
                            try { p.Set(ElementId.InvalidElementId); }
                            catch { }
                        }
                        else if (p.StorageType == StorageType.Integer)
                        {
                            try { p.Set(0); }
                            catch { }
                        }
                    }

                    // Elevation Base → Survey Point
                    // Search by name since SPOT_ELEV_BASE_PARAM
                    // doesn't exist in Revit 2023
                    if (pName.IndexOf("Elevation Base",
                            StringComparison.OrdinalIgnoreCase) >= 0
                        || pName.IndexOf("Base de elevaci",
                            StringComparison.OrdinalIgnoreCase) >= 0
                        || pName.IndexOf("Elevation Origin",
                            StringComparison.OrdinalIgnoreCase) >= 0
                        || pName.IndexOf("Origen de elevaci",
                            StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        if (p.StorageType == StorageType.Integer)
                        {
                            int before = p.AsInteger();
                            // Try each value and log for debugging
                            // UI shows: Project Base Point, Survey Point, Relative
                            // We need to find which integer = Survey Point
                            // Trying 1 first based on API docs
                            p.Set(1);
                            int after = p.AsInteger();
                            foundBase = true;
                            debugInfo.Add(
                                $"  '{typeName}': '{pName}' "
                              + $"before={before}, set=1, after={after}");
                        }
                    }
                }

                // No BuiltInParameter fallback needed - name search above handles it

                if (!foundIndicator)
                    debugInfo.Add(
                        $"  '{typeName}': NO Indicator param found. "
                      + $"RW params: {string.Join(", ", typeParamNames)}");
                if (!foundBase)
                    debugInfo.Add(
                        $"  '{typeName}': NO Base param found.");

                // ── White leader color via BuiltInParameter fallback ──
                if (whiteLeader && !foundColor)
                {
                    try
                    {
                        Parameter colorP = newType.get_Parameter(
                            BuiltInParameter.LINE_COLOR);
                        if (colorP != null && !colorP.IsReadOnly)
                        {
                            colorP.Set(16777215); // white
                            foundColor = true;
                        }
                    }
                    catch { }

                    if (!foundColor)
                        debugInfo.Add(
                            $"  '{typeName}': Could not set white leader color");
                }

                return newType.Id;
            }
            catch
            {
                return ElementId.InvalidElementId;
            }
        }

        // ════════════════════════════════════════════════════════
        // Multi-ray: find HIGHEST face on element
        // Dense grid + exact corners + small-inset corners
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

            var sampleXY = new List<XYZ>();

            // 1) 7×7 regular grid (5% to 95%)
            for (int ix = 0; ix < 7; ix++)
            {
                double px = xMin + dx * (0.05 + 0.9 * ix / 6.0);
                for (int iy = 0; iy < 7; iy++)
                {
                    double py = yMin + dy * (0.05 + 0.9 * iy / 6.0);
                    sampleXY.Add(new XYZ(px, py, 0));
                }
            }

            // 2) Exact corners (no inset — catches columns
            //    that extend to the very edge of the bbox)
            sampleXY.Add(new XYZ(xMin, yMin, 0));
            sampleXY.Add(new XYZ(xMax, yMin, 0));
            sampleXY.Add(new XYZ(xMin, yMax, 0));
            sampleXY.Add(new XYZ(xMax, yMax, 0));

            // 3) Corners with small absolute insets
            //    (15mm, 30mm, 60mm, 100mm, 150mm in feet)
            double[] absInsets = { 0.05, 0.1, 0.2, 0.33, 0.5 };
            foreach (double ins in absInsets)
            {
                sampleXY.Add(new XYZ(xMin + ins, yMin + ins, 0));
                sampleXY.Add(new XYZ(xMax - ins, yMin + ins, 0));
                sampleXY.Add(new XYZ(xMin + ins, yMax - ins, 0));
                sampleXY.Add(new XYZ(xMax - ins, yMax - ins, 0));
            }

            // 4) Edge midpoints (exact + inset)
            sampleXY.Add(new XYZ((xMin + xMax) / 2, yMin, 0));
            sampleXY.Add(new XYZ((xMin + xMax) / 2, yMax, 0));
            sampleXY.Add(new XYZ(xMin, (yMin + yMax) / 2, 0));
            sampleXY.Add(new XYZ(xMax, (yMin + yMax) / 2, 0));

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

        // ── Safe parameter setters ──────────────────────────────

        private void TrySet(Element e, BuiltInParameter bip,
            double v)
        {
            try
            {
                Parameter p = e.get_Parameter(bip);
                if (p != null && !p.IsReadOnly) p.Set(v);
            }
            catch { }
        }

        private void TrySet(Element e, BuiltInParameter bip,
            int v)
        {
            try
            {
                Parameter p = e.get_Parameter(bip);
                if (p != null && !p.IsReadOnly) p.Set(v);
            }
            catch { }
        }

        private void TrySet(Element e, BuiltInParameter bip,
            string v)
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