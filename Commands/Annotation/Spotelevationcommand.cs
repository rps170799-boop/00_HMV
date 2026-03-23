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
            // NOTE: This assumes you updated your SpotElevationWindow constructor!
            var win = new SpotElevationWindow(linkInfos, viewNames);
            if (win.ShowDialog() != true || win.Settings == null)
                return Result.Cancelled;

            SpotElevationSettings settings = win.Settings;
            double leaderOffset = settings.LeaderOffsetMm * MmToFeet;
            bool offsetX = settings.OffsetX;
            bool offsetY = settings.OffsetY;
            bool hmvStandard = settings.UseHmvStandard;
            bool createGrid = settings.CreateGrid;

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
            var skipped = new List<string>();
            var debugInfo = new List<string>();

            // Confirm which view is actually being used in the debug output
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

                    if (napHit == null)
                    {
                        skipped.Add($"{fd.Name} → No floor face hit");
                        continue;
                    }

                    Reference napRef = napHit.GetReference();
                    XYZ napPoint = napRef.GlobalPoint ?? new XYZ(hostCenter.X, hostCenter.Y, rayOrigin.Z - napHit.Proximity);
                    double bendX = napPoint.X + offX;
                    double bendY = napPoint.Y + offY;

                    if (hmvStandard)
                    {
                        var pair = new SpotPair();

                        // NTCE: strict foundation hit using Target Element ID
                        var ntceHit = FindHighestFaceInBBox(elemRI, fd.BBoxMin, fd.BBoxMax, fd.LinkInstanceId, fd.ElementId, debugInfo);

                        if (ntceHit != null)
                        {
                            Reference ntceRef = ntceHit.GetReference();
                            XYZ ntcePoint = ntceRef.GlobalPoint;
                            if (ntcePoint == null)
                            {
                                double topZ = Math.Max(fd.BBoxMax.Z, fd.BBoxMin.Z);
                                ntcePoint = new XYZ(hostCenter.X, hostCenter.Y, topZ);
                            }

                            XYZ ntceBend = new XYZ(bendX, bendY, ntcePoint.Z);
                            XYZ ntceEnd = new XYZ(bendX + offX * 0.5, bendY, ntcePoint.Z);

                            try
                            {
                                SpotDimension spotNtce = doc.Create.NewSpotElevation(view, ntceRef, ntcePoint, ntceBend, ntceEnd, ntcePoint, true);
                                if (spotNtce != null) { pair.Ntce = spotNtce; ntceSpots.Add(spotNtce); placedNtce++; }
                            }
                            catch (Exception ex) { skipped.Add($"{fd.Name} → NTCE: {ex.Message}"); }
                        }
                        else
                        {
                            skipped.Add($"{fd.Name} → NTCE: no top face found (Total rays failed)");
                        }

                        // NAP
                        XYZ napBend = new XYZ(bendX, bendY, napPoint.Z);
                        XYZ napEnd = new XYZ(bendX + offX * 0.5, bendY, napPoint.Z);

                        try
                        {
                            SpotDimension spotNap = doc.Create.NewSpotElevation(view, napRef, napPoint, napBend, napEnd, napPoint, true);
                            if (spotNap != null)
                            {
                                pair.Nap = spotNap; napSpots.Add(spotNap); placedNap++;
                                if (createGrid) gridLineData.Add(new KeyValuePair<XYZ, XYZ>(napPoint, napEnd));
                            }
                        }
                        catch (Exception ex) { skipped.Add($"{fd.Name} → NAP: {ex.Message}"); }

                        spotPairs.Add(pair);
                    }
                    else
                    {
                        // Simple mode
                        XYZ bend = new XYZ(bendX, bendY, napPoint.Z);
                        XYZ end = new XYZ(bendX + offX * 0.5, bendY, napPoint.Z);
                        try
                        {
                            SpotDimension spot = doc.Create.NewSpotElevation(view, napRef, napPoint, bend, end, napPoint, true);
                            if (spot != null)
                            {
                                placedNap++;
                                if (createGrid) gridLineData.Add(new KeyValuePair<XYZ, XYZ>(napPoint, end));
                            }
                            else skipped.Add($"{fd.Name} → null");
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

                    // Align NTCE shoulder to NAP shoulder
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
            if (hmvStandard) report += $"NTCE placed: {placedNtce}\nNAP placed: {placedNap}\n";
            else report += $"Spot elevations placed: {placedNap}\n";
            if (createGrid) report += $"Grids created: {placedGrids}\n";
            report += $"Skipped: {skipped.Count}\n";

            if (skipped.Count > 0) { report += "\n── SKIPPED ──\n"; foreach (string s in skipped) report += $"  • {s}\n"; }
            if (debugInfo.Count > 0) { report += "\n── DEBUG ──\n"; foreach (string s in debugInfo) report += $"  ▸ {s}\n"; }

            var td = new TaskDialog("HMV Tools – Spot Elevation") { MainContent = report };
            td.Show();
            return Result.Succeeded;
        }

        private ReferenceWithContext FindHighestFaceInBBox(ReferenceIntersector intersector, XYZ bbMin, XYZ bbMax, ElementId linkInstanceId, ElementId targetElementId, List<string> debugInfo)
        {
            double xMin = Math.Min(bbMin.X, bbMax.X), xMax = Math.Max(bbMin.X, bbMax.X);
            double yMin = Math.Min(bbMin.Y, bbMax.Y), yMax = Math.Max(bbMin.Y, bbMax.Y);
            double zTop = Math.Max(bbMin.Z, bbMax.Z);
            double dx = xMax - xMin, dy = yMax - yMin;

            var samples = new List<XYZ>();
            for (int ix = 0; ix < 7; ix++)
                for (int iy = 0; iy < 7; iy++)
                    samples.Add(new XYZ(xMin + dx * (0.05 + 0.9 * ix / 6.0), yMin + dy * (0.05 + 0.9 * iy / 6.0), 0));

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

            ReferenceWithContext bestHit = null;
            double bestProx = double.MaxValue;
            int totalHits = 0, linkHits = 0;

            foreach (XYZ xy in samples)
            {
                var hits = intersector.Find(new XYZ(xy.X, xy.Y, zTop + 200), XYZ.BasisZ.Negate());
                if (hits == null) continue;

                foreach (var rwc in hits)
                {
                    totalHits++;
                    Reference r = rwc.GetReference();
                    bool match = linkInstanceId != ElementId.InvalidElementId
                        ? (r.ElementId == linkInstanceId && r.LinkedElementId == targetElementId)
                        : (r.ElementId == targetElementId);

                    if (!match) continue;
                    linkHits++;
                    if (rwc.Proximity < bestProx) { bestProx = rwc.Proximity; bestHit = rwc; }
                }
            }
            debugInfo.Add($"  NTCE rays={samples.Count}, total={totalHits}, match={linkHits}, found={bestHit != null}");
            return bestHit;
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

                if (!fInd) debugInfo.Add($"  '{typeName}': NO Indicator found");
                if (!fBase) debugInfo.Add($"  '{typeName}': NO Base found");
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