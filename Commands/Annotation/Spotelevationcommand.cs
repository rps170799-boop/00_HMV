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

        private class SpotPair
        {
            public SpotDimension Ntce;
            public SpotDimension Nap;
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
                TaskDialog.Show("HMV Tools", "This command only works in plan views.");
                return Result.Cancelled;
            }
         

            // ── Collect Links ───────────────────────────────────
            var linkInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(li => li.GetLinkDocument() != null)
                .ToList();

            if (linkInstances.Count == 0)
            {
                TaskDialog.Show("HMV Tools", "No loaded Revit links found.");
                return Result.Cancelled;
            }

            var linkInfos = new List<LinkInfo>();
            for (int i = 0; i < linkInstances.Count; i++)
                linkInfos.Add(new LinkInfo { Name = linkInstances[i].Name, Index = i });

            // ── Collect Valid 3D Views for User to Choose ───────
            var valid3DViews = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .Where(v => !v.IsTemplate && !v.IsSectionBoxActive)
                .OrderBy(v => v.Name)
                .ToList();

            if (valid3DViews.Count == 0)
            {
                TaskDialog.Show("HMV Tools", "No valid 3D views found! Please ensure there is at least one 3D view without an active section box before running this tool.");
                return Result.Cancelled;
            }

            var viewNames = valid3DViews.Select(v => v.Name).ToList();

            // ── Show Window ─────────────────────────────────────
            var win = new SpotElevationWindow(linkInfos, viewNames);
            if (win.ShowDialog() != true || win.Settings == null)
                return Result.Cancelled;

            SpotElevationSettings settings = win.Settings;
            double leaderOffset = settings.LeaderOffsetMm * MmToFeet;
            bool offsetX = settings.OffsetX;
            bool offsetY = settings.OffsetY;
            bool hmvStandard = settings.UseHmvStandard;
            bool createGrid = settings.CreateGrid;
            bool hasShoulder = settings.HasShoulder; // Extract shoulder boolean

            RevitLinkInstance floorLink = linkInstances[settings.FloorLinkIndex];

            // Get the EXACT 3D view the user chose from the dropdown
            View3D view3d = valid3DViews[settings.View3DIndex];

            // ── Pick elements ───────────────────────────────────
            IList<Reference> pickedRefs;
            bool fromLink = settings.FoundationSourceIndex >= 0;

            try
            {
                pickedRefs = fromLink
                    ? uidoc.Selection.PickObjects(ObjectType.LinkedElement, "Select elements in link. Finish when done.")
                    : uidoc.Selection.PickObjects(ObjectType.Element, "Select elements. Finish when done.");
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

                    // Properly transform all 8 corners for rotated links
                    List<XYZ> corners = new List<XYZ>
                    {
                        new XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z),
                        new XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z),
                        new XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z),
                        new XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z),
                        new XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
                        new XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
                        new XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
                        new XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z)
                    };

                    var tCorners = corners.Select(c => t.OfPoint(c)).ToList();
                    bbMinH = new XYZ(tCorners.Min(c => c.X), tCorners.Min(c => c.Y), tCorners.Min(c => c.Z));
                    bbMaxH = new XYZ(tCorners.Max(c => c.X), tCorners.Max(c => c.Y), tCorners.Max(c => c.Z));
                    center = (bbMinH + bbMaxH) / 2.0;

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
                    Name = $"{elem.Category?.Name ?? "Element"}: {elem.Name} (Id {elem.Id.IntegerValue})",
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
                var cats = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_StructuralFoundation,
                    BuiltInCategory.OST_StructuralColumns,
                    BuiltInCategory.OST_Columns,
                    BuiltInCategory.OST_GenericModel,
                    BuiltInCategory.OST_StructuralFraming
                };
                elemRI = new ReferenceIntersector(
                    new ElementMulticategoryFilter(cats),
                    FindReferenceTarget.Face, view3d);
                elemRI.FindReferencesInRevitLinks = true;
            }

            // ── Process ─────────────────────────────────────────
            int placedNap = 0, placedNtce = 0, placedGrids = 0;
            int placedNapText = 0;
            ElementId topoNapTextTypeId = ElementId.InvalidElementId;
            var skipped = new List<string>();
            var debugInfo = new List<string>();

            debugInfo.Add($"Using Selected 3D View: '{view3d.Name}'");

            var ntceSpots = new List<SpotDimension>();
            var napSpots = new List<SpotDimension>();
            var spotPairs = new List<SpotPair>();
            var gridLineData = new List<KeyValuePair<XYZ, XYZ>>();

            using (Transaction tx = new Transaction(doc, "HMV – Spot Elevations"))
            {
                tx.Start();

                foreach (var fd in foundations)
                {
                    XYZ hostCenter = fd.HostCenter;
                    double offX = offsetX ? leaderOffset : 0;
                    double offY = offsetY ? leaderOffset : 0;

                    // NAP: Floor hit
                    XYZ rayOrigin = new XYZ(hostCenter.X, hostCenter.Y, hostCenter.Z + 200);
                    var floorHits = floorRI.Find(rayOrigin, XYZ.BasisZ.Negate());
                    var napHit = FindFirstHitOnLink(floorHits, floorLink.Id);




                    XYZ topoTextPoint = null;
                    if (napHit == null)
                        topoTextPoint = FindTopoMeshHit(floorLink, hostCenter.X, hostCenter.Y, debugInfo);

                    if (napHit == null && topoTextPoint == null)
                    {
                        skipped.Add($"{fd.Name} → No floor or topo hit");
                        continue;
                    }

                    Reference napRef;
                    XYZ napPoint;
                    bool isTopoText;
                    if (napHit != null)
                    {
                        napRef = napHit.GetReference();
                        napPoint = napRef.GlobalPoint ?? new XYZ(hostCenter.X, hostCenter.Y, rayOrigin.Z - napHit.Proximity);
                        isTopoText = false;
                    }
                    else
                    {
                        napRef = null;
                        napPoint = topoTextPoint;
                        isTopoText = true;
                    }

                    if (hmvStandard)
                    {
                        var pair = new SpotPair();
                        var ntceHit = FindHighestFaceInBBox(elemRI, fd.BBoxMin, fd.BBoxMax, fd.LinkInstanceId, fd.ElementId, debugInfo);

                        if (ntceHit != null)
                        {
                            Reference ntceRef = ntceHit.GetReference();
                            XYZ ntcePoint = ntceRef.GlobalPoint ?? new XYZ(hostCenter.X, hostCenter.Y, Math.Max(fd.BBoxMax.Z, fd.BBoxMin.Z));

                            XYZ napBend, napEnd, ntceBend, ntceEnd;

                            if (hasShoulder)
                            {
                                // ORIGINAL LOGIC: Horizontal alignment
                                double bendX = napPoint.X + offX;
                                double bendY = napPoint.Y + offY;

                                ntceBend = new XYZ(bendX, bendY, ntcePoint.Z);
                                ntceEnd = new XYZ(bendX + offX * 0.5, bendY, ntcePoint.Z);

                                napBend = new XYZ(bendX, bendY, napPoint.Z);
                                napEnd = new XYZ(bendX + offX * 0.5, bendY, napPoint.Z);
                            }
                            else
                            {
                                // VERTICAL / DIAGONAL STACK LOGIC (Simulating "No Shoulder")
                                // We make the shoulder perfectly collinear with the leader vector 
                                // so it visually becomes one perfectly straight line.

                                double gapX = 350 * MmToFeet; // Shift NTCE left
                                double gapY = 350 * MmToFeet; // Shift NTCE up

                                // 1. Set the final text destination (End point)
                                napEnd = new XYZ(hostCenter.X + offX, hostCenter.Y + offY, napPoint.Z);

                                if (offsetY && !offsetX)
                                {
                                    // Stack vertically, NTCE shifted left and up
                                    ntceEnd = new XYZ(hostCenter.X + offX - gapX, hostCenter.Y + offY + gapY, ntcePoint.Z);
                                }
                                else
                                {
                                    // Default fallback
                                    ntceEnd = new XYZ(hostCenter.X + offX + gapX, hostCenter.Y + offY + gapY, ntcePoint.Z);
                                }

                                // 2. The Hack: Place the bend exactly halfway between origin and end.
                                // This gives the shoulder a non-zero length, but forces it to point
                                // perfectly straight at the element, hiding the horizontal break.
                                napBend = napPoint + (napEnd - napPoint) * 0.5;
                                ntceBend = ntcePoint + (ntceEnd - ntcePoint) * 0.5;
                            }

                            try
                            {
                                SpotDimension spotNtce = doc.Create.NewSpotElevation(view, ntceRef, ntcePoint, ntceBend, ntceEnd, ntcePoint, true);
                                if (spotNtce != null) { pair.Ntce = spotNtce; ntceSpots.Add(spotNtce); placedNtce++; }
                            }
                            catch (Exception ex) { skipped.Add($"{fd.Name} → NTCE: {ex.Message}"); }

                            try
                            {
                                if (isTopoText)
                                {
                                    if (topoNapTextTypeId == ElementId.InvalidElementId)
                                        topoNapTextTypeId = FindOrCreateTopoNapTextType(doc, "HMV_General_2mm Arial");

                                    if (topoNapTextTypeId != ElementId.InvalidElementId)
                                    {
                                        string txt = FormatNapText(doc, napPoint);
                                        TextNote tn = TextNote.Create(doc, view.Id, napEnd, txt, topoNapTextTypeId);
                                        if (tn != null)
                                        {
                                            placedNap++;
                                            placedNapText++;
                                            if (createGrid) gridLineData.Add(new KeyValuePair<XYZ, XYZ>(napPoint, napEnd));
                                        }
                                    }
                                }
                                else
                                {
                                    SpotDimension spotNap = doc.Create.NewSpotElevation(view, napRef, napPoint, napBend, napEnd, napPoint, true);
                                    if (spotNap != null)
                                    {
                                        pair.Nap = spotNap; napSpots.Add(spotNap); placedNap++;
                                        if (createGrid) gridLineData.Add(new KeyValuePair<XYZ, XYZ>(napPoint, napEnd));
                                    }
                                }
                            }
                            catch (Exception ex) { skipped.Add($"{fd.Name} → NAP: {ex.Message}"); }
                        }
                        else
                        {
                            skipped.Add($"{fd.Name} → NTCE: no top face found (Total rays failed)");
                        }

                        spotPairs.Add(pair);
                    }
                    else
                    {
                        // Simple mode
                        XYZ bend, end;
                        if (hasShoulder)
                        {
                            bend = new XYZ(napPoint.X + offX, napPoint.Y + offY, napPoint.Z);
                            end = new XYZ(napPoint.X + offX * 1.5, napPoint.Y + offY, napPoint.Z);
                        }
                        else
                        {
                            // Collinear Hack for simple mode too
                            end = new XYZ(hostCenter.X + offX, hostCenter.Y + offY, napPoint.Z);
                            bend = napPoint + (end - napPoint) * 0.5;
                        }

                        try
                        {
                            if (isTopoText)
                            {
                                if (topoNapTextTypeId == ElementId.InvalidElementId)
                                    topoNapTextTypeId = FindOrCreateTopoNapTextType(doc, "HMV_General_2mm Arial");

                                if (topoNapTextTypeId != ElementId.InvalidElementId)
                                {
                                    string txt = FormatNapText(doc, napPoint);
                                    TextNote tn = TextNote.Create(doc, view.Id, end, txt, topoNapTextTypeId);
                                    if (tn != null)
                                    {
                                        placedNap++;
                                        placedNapText++;
                                        if (createGrid) gridLineData.Add(new KeyValuePair<XYZ, XYZ>(napPoint, end));
                                    }
                                }
                            }
                            else
                            {
                                SpotDimension spot = doc.Create.NewSpotElevation(view, napRef, napPoint, bend, end, napPoint, true);
                                if (spot != null)
                                {
                                    placedNap++;
                                    if (createGrid) gridLineData.Add(new KeyValuePair<XYZ, XYZ>(napPoint, end));
                                }
                                else skipped.Add($"{fd.Name} → null");
                            }
                        }
                        catch (Exception ex) { skipped.Add($"{fd.Name} → {ex.Message}"); }
                    }
                }

                if (hmvStandard)
                {
                    SpotDimension anySpot = ntceSpots.FirstOrDefault() ?? napSpots.FirstOrDefault();
                    if (anySpot != null)
                    {
                        ElementId baseTypeId = anySpot.GetTypeId();
                        ElementType baseType = doc.GetElement(baseTypeId) as ElementType;

                        if (baseType != null)
                        {
                            ElementId ntceTypeId = FindOrCreateType(doc, baseType, NtceTypeName, "N.T.C.E.", false, true, debugInfo);
                            ElementId napTypeId = FindOrCreateType(doc, baseType, NapTypeName, "N.A.P.", true, false, debugInfo);

                            if (ntceTypeId != ElementId.InvalidElementId)
                                foreach (var s in ntceSpots) try { s.ChangeTypeId(ntceTypeId); } catch { }

                            if (napTypeId != ElementId.InvalidElementId)
                                foreach (var s in napSpots) try { s.ChangeTypeId(napTypeId); } catch { }
                        }
                    }

                    // Align NTCE shoulder to NAP shoulder (ONLY if HasShoulder is TRUE)
                    if (hasShoulder)
                    {
                        foreach (var pair in spotPairs)
                        {
                            if (pair.Ntce == null || pair.Nap == null) continue;
                            try
                            {
                                XYZ napShoulder = pair.Nap.LeaderShoulderPosition;
                                if (napShoulder != null) pair.Ntce.LeaderShoulderPosition = new XYZ(napShoulder.X, napShoulder.Y, napShoulder.Z);
                            }
                            catch { }
                        }
                    }
                }

                if (createGrid && gridLineData.Count > 0)
                {
                    double gridTol = 1.0 * MmToFeet;
                    if (offsetX)
                    {
                        var yGroups = GroupByCoordinate(gridLineData, p => p.Key.Y, gridTol);
                        foreach (var g in yGroups)
                        {
                            double y = g.Key;
                            double minX = g.Value.Min(p => Math.Min(p.Key.X, p.Value.X));
                            double maxX = g.Value.Max(p => Math.Max(p.Key.X, p.Value.X));
                            if (Math.Abs(maxX - minX) < gridTol) continue;
                            try { Grid grid = Grid.Create(doc, Line.CreateBound(new XYZ(minX, y, 0), new XYZ(maxX, y, 0))); if (grid != null) { grid.Pinned = false; placedGrids++; } }
                            catch { }
                        }
                    }
                    if (offsetY)
                    {
                        var xGroups = GroupByCoordinate(gridLineData, p => p.Key.X, gridTol);
                        foreach (var g in xGroups)
                        {
                            double x = g.Key;
                            double minY = g.Value.Min(p => Math.Min(p.Key.Y, p.Value.Y));
                            double maxY = g.Value.Max(p => Math.Max(p.Key.Y, p.Value.Y));
                            if (Math.Abs(maxY - minY) < gridTol) continue;
                            try { Grid grid = Grid.Create(doc, Line.CreateBound(new XYZ(x, minY, 0), new XYZ(x, maxY, 0))); if (grid != null) { grid.Pinned = false; placedGrids++; } }
                            catch { }
                        }
                    }
                }
                tx.Commit();
            }

            // Report
            string report = "SPOT ELEVATION SUMMARY\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
                + $"Elements picked: {foundations.Count}\n";
            if (hmvStandard) report += $"NTCE placed: {placedNtce}\nNAP placed: {placedNap}" + (placedNapText > 0 ? $" ({placedNapText} as text on topo)" : "") + "\n";
            else report += $"Spot elevations placed: {placedNap}" + (placedNapText > 0 ? $" ({placedNapText} as text on topo)" : "") + "\n";
            if (createGrid) report += $"Grids created: {placedGrids}\n";
            report += $"Skipped: {skipped.Count}\n";

            if (skipped.Count > 0) { report += "\n── SKIPPED ──\n"; foreach (string s in skipped) report += $"  • {s}\n"; }
            if (debugInfo.Count > 0) { report += "\n── DEBUG ──\n"; foreach (string s in debugInfo) report += $"  ▸ {s}\n"; }

            var td = new TaskDialog("HMV Tools – Spot Elevation") { MainContent = report };
            td.Show();
            return Result.Succeeded;
        }

        /// <summary>
        /// Raycasts downward across a dense sample grid within the element's
        /// bounding box. Two-step filter:
        ///   1) Group hit Z values by 5 mm tolerance, discard groups with
        ///      less than 5% of total rays (noise: bolts, rebar tips).
        ///   2) Among surviving groups, pick the HIGHEST Z — this selects
        ///      the pedestal top over the footing pad.
        /// </summary>
        private ReferenceWithContext FindHighestFaceInBBox(ReferenceIntersector intersector, XYZ bbMin, XYZ bbMax, ElementId linkInstanceId, ElementId targetElementId, List<string> debugInfo)
        {
            double xMin = Math.Min(bbMin.X, bbMax.X), xMax = Math.Max(bbMin.X, bbMax.X);
            double yMin = Math.Min(bbMin.Y, bbMax.Y), yMax = Math.Max(bbMin.Y, bbMax.Y);
            double zTop = Math.Max(bbMin.Z, bbMax.Z);
            double dx = xMax - xMin, dy = yMax - yMin;

            var samples = new List<XYZ>();
            for (int ix = 0; ix < 11; ix++)
                for (int iy = 0; iy < 11; iy++)
                    samples.Add(new XYZ(xMin + dx * (0.05 + 0.9 * ix / 10.0), yMin + dy * (0.05 + 0.9 * iy / 10.0), 0));

            samples.Add(new XYZ(xMin, yMin, 0)); samples.Add(new XYZ(xMax, yMin, 0));
            samples.Add(new XYZ(xMin, yMax, 0)); samples.Add(new XYZ(xMax, yMax, 0));

            double[] ins = { 0.05, 0.1, 0.2, 0.33, 0.5 };
            foreach (double i in ins)
            {
                samples.Add(new XYZ(xMin + i, yMin + i, 0)); samples.Add(new XYZ(xMax - i, yMin + i, 0));
                samples.Add(new XYZ(xMin + i, yMax - i, 0)); samples.Add(new XYZ(xMax - i, yMax - i, 0));
            }

            samples.Add(new XYZ((xMin + xMax) / 2, yMin, 0)); samples.Add(new XYZ((xMin + xMax) / 2, yMax, 0));
            samples.Add(new XYZ(xMin, (yMin + yMax) / 2, 0)); samples.Add(new XYZ(xMax, (yMin + yMax) / 2, 0));

            double rayOriginZ = zTop + 200;
            int totalHits = 0, matchHits = 0;

            // Collect all matching hits with their Z values
            var hitData = new List<KeyValuePair<double, ReferenceWithContext>>();

            foreach (XYZ xy in samples)
            {
                var hits = intersector.Find(new XYZ(xy.X, xy.Y, rayOriginZ), XYZ.BasisZ.Negate());
                if (hits == null) continue;

                // Per ray, take only the first (closest) hit matching our element
                double bestProxThisRay = double.MaxValue;
                double bestZThisRay = 0;
                ReferenceWithContext bestRwcThisRay = null;

                foreach (var rwc in hits)
                {
                    totalHits++;
                    Reference r = rwc.GetReference();
                    bool match = linkInstanceId != ElementId.InvalidElementId
                        ? (r.ElementId == linkInstanceId && r.LinkedElementId == targetElementId)
                        : (r.ElementId == targetElementId);

                    if (!match) continue;
                    matchHits++;

                    if (rwc.Proximity < bestProxThisRay)
                    {
                        bestProxThisRay = rwc.Proximity;
                        bestZThisRay = r.GlobalPoint != null
                            ? r.GlobalPoint.Z
                            : (rayOriginZ - rwc.Proximity);
                        bestRwcThisRay = rwc;
                    }
                }

                if (bestRwcThisRay != null)
                    hitData.Add(new KeyValuePair<double, ReferenceWithContext>(bestZThisRay, bestRwcThisRay));
            }

            debugInfo.Add($"  NTCE rays={samples.Count}, total={totalHits}, match={matchHits}, zValues={hitData.Count}");

            if (hitData.Count == 0) return null;

            // Two-step filter:
            //   Step 1 — Group Z values by tolerance (5 mm ≈ 0.017 ft),
            //            discard groups with < 5% of total rays (noise: bolts, rebar).
            //   Step 2 — Among surviving groups, pick the HIGHEST Z.
            //            This selects the pedestal top over the footing pad.
            double groupTol = 0.017; // ~5 mm in feet
            var groups = new List<KeyValuePair<double, List<ReferenceWithContext>>>(); // avgZ, hits in group

            foreach (var entry in hitData)
            {
                double z = entry.Key;
                bool merged = false;
                for (int gi = 0; gi < groups.Count; gi++)
                {
                    if (Math.Abs(z - groups[gi].Key) < groupTol)
                    {
                        // Running average + add hit to group
                        var grp = groups[gi];
                        int newCount = grp.Value.Count + 1;
                        double newAvg = grp.Key + (z - grp.Key) / newCount;
                        grp.Value.Add(entry.Value);
                        groups[gi] = new KeyValuePair<double, List<ReferenceWithContext>>(newAvg, grp.Value);
                        merged = true;
                        break;
                    }
                }
                if (!merged)
                    groups.Add(new KeyValuePair<double, List<ReferenceWithContext>>(
                        z, new List<ReferenceWithContext> { entry.Value }));
            }

            // Step 1: Discard noise — groups with < 5% of total ray hits
            int minVotes = Math.Max(1, (int)(hitData.Count * 0.05));
            var significant = groups.Where(g => g.Value.Count >= minVotes).ToList();

            // Fallback: if all groups got filtered, keep the one with most votes
            if (significant.Count == 0)
                significant = new List<KeyValuePair<double, List<ReferenceWithContext>>>
                    { groups.OrderByDescending(g => g.Value.Count).First() };

            // Debug: log all groups so we can verify
            foreach (var g in groups.OrderByDescending(g => g.Key))
                debugInfo.Add($"  NTCE-group: Z={g.Key * 0.3048:F4}m, votes={g.Value.Count}/{hitData.Count}{(g.Value.Count < minVotes ? " [NOISE]" : "")}");

            // Step 2: Among significant groups, pick the HIGHEST Z
            var winner = significant.OrderByDescending(g => g.Key).First();
            debugInfo.Add($"  NTCE-winner: {winner.Key * 0.3048:F4}m, votes={winner.Value.Count}/{hitData.Count}, threshold={minVotes}");

            // Return the hit with smallest proximity within the winning group
            return winner.Value.OrderBy(rwc => rwc.Proximity).First();
        }

        private XYZ FindTopoMeshHit(RevitLinkInstance link, double xHost, double yHost, List<string> debugInfo)
        {
            Document linkDoc = link?.GetLinkDocument();
            if (linkDoc == null) return null;

            Transform linkT = link.GetTotalTransform();
            Transform invT = linkT.Inverse;
            XYZ probeLocal = invT.OfPoint(new XYZ(xHost, yHost, 0));
            double xL = probeLocal.X, yL = probeLocal.Y;

            var collector = new FilteredElementCollector(linkDoc)
                .OfCategory(BuiltInCategory.OST_Topography)
                .WhereElementIsNotElementType();

            var opt = new Options();
            double bestZLocal = double.NegativeInfinity;
            int triCount = 0, hitCount = 0;
            bool found = false;

            foreach (Element topo in collector)
            {
                GeometryElement geom = topo.get_Geometry(opt);
                if (geom == null) continue;
                foreach (GeometryObject gObj in geom)
                {
                    if (!(gObj is Mesh mesh)) continue;
                    for (int i = 0; i < mesh.NumTriangles; i++)
                    {
                        triCount++;
                        MeshTriangle tri = mesh.get_Triangle(i);
                        XYZ v0 = tri.get_Vertex(0);
                        XYZ v1 = tri.get_Vertex(1);
                        XYZ v2 = tri.get_Vertex(2);

                        double denom = (v1.Y - v2.Y) * (v0.X - v2.X) + (v2.X - v1.X) * (v0.Y - v2.Y);
                        if (Math.Abs(denom) < 1e-12) continue;
                        double a = ((v1.Y - v2.Y) * (xL - v2.X) + (v2.X - v1.X) * (yL - v2.Y)) / denom;
                        double b = ((v2.Y - v0.Y) * (xL - v2.X) + (v0.X - v2.X) * (yL - v2.Y)) / denom;
                        double c = 1.0 - a - b;

                        if (a < -1e-9 || b < -1e-9 || c < -1e-9) continue;
                        hitCount++;

                        double zHit = a * v0.Z + b * v1.Z + c * v2.Z;
                        if (zHit > bestZLocal) { bestZLocal = zHit; found = true; }
                    }
                }
            }

            debugInfo?.Add($"  topo mesh: tris={triCount}, in-XY={hitCount}, found={found}");
            if (!found) return null;

            return linkT.OfPoint(new XYZ(xL, yL, bestZLocal));
        }

        private ElementId FindOrCreateTopoNapTextType(Document doc, string typeName)
        {
            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .FirstOrDefault(t => t.Name.Equals(typeName, StringComparison.OrdinalIgnoreCase));
            if (existing != null) return existing.Id;

            var baseType = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .FirstOrDefault();
            if (baseType == null) return ElementId.InvalidElementId;

            try
            {
                TextNoteType nt = baseType.Duplicate(typeName) as TextNoteType;
                if (nt == null) return ElementId.InvalidElementId;
                TrySet(nt, BuiltInParameter.TEXT_SIZE, 2.0 * MmToFeet);
                TrySet(nt, BuiltInParameter.TEXT_FONT, "Arial");
                TrySet(nt, BuiltInParameter.TEXT_BACKGROUND, 0);
                return nt.Id;
            }
            catch { return ElementId.InvalidElementId; }
        }

        private string FormatNapText(Document doc, XYZ hostPoint)
        {
            Transform internalToSurvey = doc.ActiveProjectLocation.GetTransform().Inverse;
            XYZ survey = internalToSurvey.OfPoint(hostPoint);
            double zMeters = UnitUtils.ConvertFromInternalUnits(survey.Z, UnitTypeId.Meters);
            return "N.A.P. " + zMeters.ToString("0.00000", System.Globalization.CultureInfo.InvariantCulture);
        }

        private ReferenceWithContext FindFirstHitOnLink(IList<ReferenceWithContext> hits, ElementId linkId)
        {
            if (hits == null) return null;
            ReferenceWithContext best = null;
            double bestProx = double.MaxValue;
            foreach (var rwc in hits)
            {
                if (rwc.GetReference().ElementId != linkId) continue;
                if (rwc.Proximity < bestProx) { bestProx = rwc.Proximity; best = rwc; }
            }
            return best;
        }

        private ElementId FindOrCreateType(Document doc, ElementType baseType, string typeName, string prefix, bool textBelowLeader, bool whiteLeader, List<string> debugInfo)
        {
            foreach (Element e in new FilteredElementCollector(doc).WhereElementIsElementType().ToElements())
                if (e.Name == typeName && e.GetType().Name == baseType.GetType().Name) return e.Id;

            try
            {
                ElementType nt = baseType.Duplicate(typeName);
                if (nt == null) return ElementId.InvalidElementId;
                TrySet(nt, BuiltInParameter.TEXT_SIZE, 2.5 * MmToFeet);
                TrySet(nt, BuiltInParameter.TEXT_FONT, "Arial");
                TrySet(nt, BuiltInParameter.TEXT_STYLE_BOLD, 0);
                TrySet(nt, BuiltInParameter.TEXT_BACKGROUND, 0);
                TrySet(nt, BuiltInParameter.TEXT_WIDTH_SCALE, 1.0);
                TrySet(nt, BuiltInParameter.SPOT_ELEV_TEXT_LOCATION, textBelowLeader ? 1 : 0);
                TrySet(nt, BuiltInParameter.SPOT_ELEV_DISPLAY_ELEVATIONS, 0);

                bool fInd = false, fBase = false, fColor = false;
                foreach (Parameter p in nt.Parameters)
                {
                    string pn = p.Definition?.Name;
                    if (pn == null || p.IsReadOnly) continue;

                    if (pn.Equals("Elevation Indicator", StringComparison.OrdinalIgnoreCase) || pn.Equals("Indicador de elevación", StringComparison.OrdinalIgnoreCase) || pn.Equals("Indicador de elevacion", StringComparison.OrdinalIgnoreCase))
                    { if (p.StorageType == StorageType.String) { p.Set(prefix); fInd = true; } }
                    if (pn.IndexOf("Bottom Indicator", StringComparison.OrdinalIgnoreCase) >= 0 || pn.IndexOf("Top Indicator", StringComparison.OrdinalIgnoreCase) >= 0 || pn.IndexOf("Indicador inferior", StringComparison.OrdinalIgnoreCase) >= 0 || pn.IndexOf("Indicador superior", StringComparison.OrdinalIgnoreCase) >= 0)
                    { if (p.StorageType == StorageType.String) p.Set(""); }
                    if (whiteLeader && pn.Equals("Color", StringComparison.OrdinalIgnoreCase))
                    { if (p.StorageType == StorageType.Integer) try { p.Set(16777215); fColor = true; } catch { } }
                    if (whiteLeader && (pn.IndexOf("Leader Arrowhead", StringComparison.OrdinalIgnoreCase) >= 0 || pn.IndexOf("Punta de flecha", StringComparison.OrdinalIgnoreCase) >= 0))
                    { if (p.StorageType == StorageType.ElementId) try { p.Set(ElementId.InvalidElementId); } catch { } }
                    if (pn.IndexOf("Elevation Base", StringComparison.OrdinalIgnoreCase) >= 0 || pn.IndexOf("Base de elevaci", StringComparison.OrdinalIgnoreCase) >= 0 || pn.IndexOf("Elevation Origin", StringComparison.OrdinalIgnoreCase) >= 0 || pn.IndexOf("Origen de elevaci", StringComparison.OrdinalIgnoreCase) >= 0)
                    { if (p.StorageType == StorageType.Integer) { p.Set(1); fBase = true; } }
                }
                if (whiteLeader && !fColor) try { var cp = nt.get_Parameter(BuiltInParameter.LINE_COLOR); if (cp != null && !cp.IsReadOnly) cp.Set(16777215); } catch { }

                if (!fInd) debugInfo.Add($"  '{typeName}': NO Indicator");
                if (!fBase) debugInfo.Add($"  '{typeName}': NO Base");

                return nt.Id;
            }
            catch { return ElementId.InvalidElementId; }
        }

        private Dictionary<double, List<KeyValuePair<XYZ, XYZ>>> GroupByCoordinate(List<KeyValuePair<XYZ, XYZ>> data, Func<KeyValuePair<XYZ, XYZ>, double> sel, double tol)
        {
            var g = new Dictionary<double, List<KeyValuePair<XYZ, XYZ>>>();
            foreach (var p in data)
            {
                double v = sel(p); bool f = false;
                foreach (double k in g.Keys.ToList()) if (Math.Abs(v - k) < tol) { g[k].Add(p); f = true; break; }
                if (!f) g[v] = new List<KeyValuePair<XYZ, XYZ>> { p };
            }
            return g;
        }

        private void TrySet(Element e, BuiltInParameter b, double v) { try { var p = e.get_Parameter(b); if (p != null && !p.IsReadOnly) p.Set(v); } catch { } }
        private void TrySet(Element e, BuiltInParameter b, int v) { try { var p = e.get_Parameter(b); if (p != null && !p.IsReadOnly) p.Set(v); } catch { } }
        private void TrySet(Element e, BuiltInParameter b, string v) { try { var p = e.get_Parameter(b); if (p != null && !p.IsReadOnly) p.Set(v); } catch { } }
    }
}