using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using IxMilia.Dxf;
using IxMilia.Dxf.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HMVTools
{
    // ═══════════════════════════════════════════════════════════════════
    //  CONFIGURATION
    // ═══════════════════════════════════════════════════════════════════

    public class ElectricalRefresherConfig
    {
        // Parameters on adaptive point
        public string CnxNumberParam    { get; set; } = "HMV_CFI_CONEXI\u00d3N";
        public string EquipoInicialParam { get; set; } = "HMV_CFI_EQUIPO INICIAL";
        public string ConnectedParam    { get; set; } = "HMV_CFI_CONNECTED";

        // Parameters on FlexPipe
        public string FlexCnxNumberParam    { get; set; } = "HMV_CFI_CONEXI\u00d3N";
        public string FlexEquipoInicialParam { get; set; } = "HMV_CFI_EQUIPO INICIAL";
        public string FlexEquipoFinalParam   { get; set; } = "HMV_CFI_EQUIPO FINAL";

        // Types
        public string FlexPipeTypeKey        { get; set; }
        public string ElectricalSystemTypeKey { get; set; }

        // DXF folder
        public string DxfFolder { get; set; }
    }

    // ═══════════════════════════════════════════════════════════════════
    //  COMMAND  — ribbon entry point
    // ═══════════════════════════════════════════════════════════════════

    [Transaction(TransactionMode.Manual)]
    public class ElectricalRefresherCommand : IExternalCommand
    {
        private static ElectricalRefresherPickHandler _pickHandler = null;
        private static ExternalEvent                  _pickEvent   = null;
        private static ElectricalRefresherHandler     _runHandler  = null;
        private static ExternalEvent                  _runEvent    = null;
        private static ElectricalRefresherWindow      _window      = null;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (_window != null && _window.IsLoaded)
            {
                _window.Focus();
                return Result.Succeeded;
            }

            _pickHandler = new ElectricalRefresherPickHandler();
            _pickEvent   = ExternalEvent.Create(_pickHandler);
            _runHandler  = new ElectricalRefresherHandler();
            _runEvent    = ExternalEvent.Create(_runHandler);

            _window = new ElectricalRefresherWindow(
                commandData.Application,
                _pickHandler, _pickEvent,
                _runHandler,  _runEvent);

            var helper = new System.Windows.Interop.WindowInteropHelper(_window);
            helper.Owner = commandData.Application.MainWindowHandle;

            _window.Show();
            return Result.Succeeded;
        }

        public static void ClearWindow() { _window = null; }
    }

    // ═══════════════════════════════════════════════════════════════════
    //  PICK HANDLER
    // ═══════════════════════════════════════════════════════════════════

    public class ElectricalRefresherPickHandler : IExternalEventHandler
    {
        public ElectricalRefresherWindow UI { get; set; }

        public string GetName() => "ElectricalRefresherPickHandler";

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;

            try
            {
                IList<Reference> refs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    new AdaptiveComponentFilter(),
                    "Select Adaptive Points — click Finish when done");

                if (refs == null || refs.Count == 0)
                {
                    UI?.SetStatus("No elements selected.");
                    UI?.RestoreWindow();
                    return;
                }

                UI?.OnPointsPicked(refs.Select(r => r.ElementId).ToList());
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                UI?.SetStatus("Selection cancelled.");
                UI?.RestoreWindow();
            }
            catch (Exception ex)
            {
                UI?.SetStatus("Pick error: " + ex.Message);
                UI?.RestoreWindow();
            }
        }
    }

    // ═══════════════════════════════════════════════════════════════════
    //  MAIN HANDLER
    // ═══════════════════════════════════════════════════════════════════

    public class ElectricalRefresherHandler : IExternalEventHandler
    {
        public ElectricalRefresherWindow UI     { get; set; }
        public List<ElementId>           PointIds { get; set; }
        public ElectricalRefresherConfig Config  { get; set; }
        public bool                      IsReset { get; set; }

        public string GetName() => "ElectricalRefresherHandler";

        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;

            if (PointIds == null || PointIds.Count == 0 || Config == null)
            {
                UI?.Log("Error: No points or configuration.");
                return;
            }

            if (IsReset)
                ExecuteReset(doc);
            else
                ExecuteRefresh(doc);
        }

        // ─────────────────────────────────────────────────────────────
        //  RESET MODE — delete FlexPipes whose CNX_NUMBER matches the
        //               selected adaptive points
        // ─────────────────────────────────────────────────────────────
        private void ExecuteReset(Document doc)
        {
            UI?.Log("── RESET MODE ──────────────────────────────────");

            // Collect connection numbers from selected points
            var cnxNumbers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ElementId id in PointIds)
            {
                FamilyInstance fi = doc.GetElement(id) as FamilyInstance;
                if (fi == null) continue;
                string cnx = GetStringParam(fi, Config.CnxNumberParam);
                if (!string.IsNullOrWhiteSpace(cnx)) cnxNumbers.Add(cnx);
            }

            if (cnxNumbers.Count == 0)
            {
                UI?.Log("Warning: No connection numbers found on selected points.");
                return;
            }

            // Collect FlexPipes matching those numbers
            var flexPipesToDelete = new FilteredElementCollector(doc)
                .OfClass(typeof(FlexPipe))
                .Cast<FlexPipe>()
                .Where(fp =>
                {
                    string v = GetStringParam(fp, Config.FlexCnxNumberParam);
                    return !string.IsNullOrWhiteSpace(v) && cnxNumbers.Contains(v);
                })
                .Select(fp => fp.Id)
                .ToList();

            using (Transaction trans = new Transaction(doc, "HMV - Reset Electrical Connections"))
            {
                trans.Start();

                FailureHandlingOptions fho = trans.GetFailureHandlingOptions();
                fho.SetFailuresPreprocessor(new WarningSuppressor());
                trans.SetFailureHandlingOptions(fho);

                int deleted = 0;
                foreach (ElementId fpId in flexPipesToDelete)
                {
                    try { doc.Delete(fpId); deleted++; }
                    catch (Exception ex) { UI?.Log($"  Could not delete FlexPipe {fpId}: {ex.Message}"); }
                }

                // Reset HMV_CFI_CONNECTED = 0 on the points
                int reset = 0;
                foreach (ElementId id in PointIds)
                {
                    FamilyInstance fi = doc.GetElement(id) as FamilyInstance;
                    if (fi == null) continue;
                    Parameter p = fi.LookupParameter(Config.ConnectedParam);
                    if (p != null && !p.IsReadOnly && p.StorageType == StorageType.Integer)
                    { p.Set(0); reset++; }
                }

                trans.Commit();
                UI?.Log($"✔ Deleted {deleted} FlexPipe(s). Reset flag on {reset} point(s).");
            }
        }

        // ─────────────────────────────────────────────────────────────
        //  REFRESH MODE — pair points by CNX_NUMBER and create FlexPipes
        // ─────────────────────────────────────────────────────────────
        private void ExecuteRefresh(Document doc)
        {
            UI?.Log("── REFRESH MODE ─────────────────────────────────");

            if (string.IsNullOrWhiteSpace(Config.DxfFolder) || !Directory.Exists(Config.DxfFolder))
            {
                UI?.Log("Error: DXF folder not found — check configuration.");
                return;
            }

            // ── Level ──────────────────────────────────────────────────────────
            Level defaultLevel = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => l.Elevation)
                .FirstOrDefault();

            if (defaultLevel == null) { UI?.Log("Error: No Level found in project."); return; }

            // ── Resolve FlexPipe / System types ────────────────────────────────
            FlexPipeType flexType = FindFlexPipeType(doc, Config.FlexPipeTypeKey);
            if (flexType == null) { UI?.Log($"Error: FlexPipe type '{Config.FlexPipeTypeKey}' not found."); return; }

            ElementId systemTypeId = FindElectricalSystemTypeId(doc, Config.ElectricalSystemTypeKey);
            if (systemTypeId == null) { UI?.Log($"Error: System type '{Config.ElectricalSystemTypeKey}' not found."); return; }

            // ── Build DXF map: CNX_NUMBER → file path ──────────────────────────
            Dictionary<string, string> dxfMap = BuildDxfMap(Config.DxfFolder);

            // ── Collect adaptive points with their CNX_NUMBER ──────────────────
            var pointData = new List<(ElementId id, FamilyInstance fi, string cnx, string equipoId, XYZ pos)>();

            foreach (ElementId id in PointIds)
            {
                FamilyInstance fi = doc.GetElement(id) as FamilyInstance;
                if (fi == null) continue;

                string cnx = GetStringParam(fi, Config.CnxNumberParam);
                if (string.IsNullOrWhiteSpace(cnx)) { UI?.Log($"  Skip {id}: no {Config.CnxNumberParam}."); continue; }

                // Only process points where IS CONNECTED = false (0)
                Parameter connParam = fi.LookupParameter(Config.ConnectedParam);
                if (connParam != null && connParam.StorageType == StorageType.Integer && connParam.AsInteger() == 1)
                {
                    UI?.Log($"  Skip {id} [{cnx}]: already connected.");
                    continue;
                }

                string equipoId = GetStringParam(fi, Config.EquipoInicialParam);
                XYZ pos = GetAdaptivePosition(doc, fi);
                pointData.Add((id, fi, cnx, equipoId, pos));
            }

            // ── Group by CNX_NUMBER ────────────────────────────────────────────
            var groups = pointData
                .GroupBy(p => p.cnx, StringComparer.OrdinalIgnoreCase)
                .ToList();

            int created = 0;
            int skipped = 0;

            foreach (var group in groups)
            {
                var pts = group.ToList();
                string cnx = group.Key;

                if (pts.Count < 2)
                {
                    UI?.Log($"  [{cnx}] Skipped: only {pts.Count} point(s) found.");
                    skipped++;
                    continue;
                }

                // Find optimal pairs — filters same-equipment pairs, minimises distance
                var pairIndices = FindOptimalPairIndices(pts);

                if (pairIndices.Count == 0)
                {
                    UI?.Log($"  [{cnx}] Skipped: {pts.Count} point(s) but no valid pairs (all share the same equipment).");
                    skipped++;
                    continue;
                }

                if (pts.Count > 2)
                    UI?.Log($"  [{cnx}] {pts.Count} points → {pairIndices.Count} pair(s) to connect.");

                // ── Find DXF file (shared by all pairs in this group) ──────────
                if (!dxfMap.TryGetValue(cnx, out string dxfPath))
                {
                    UI?.Log($"  [{cnx}] Warning: no DXF file found for key '{cnx}'. Skipped.");
                    skipped += pairIndices.Count;
                    continue;
                }

                // Pre-parse DXF once — reused for every pair in this group
                double groupDxfDx = 0;
                double groupDxfDy = 0;
                List<XYZ> groupRawPoints = null;

                try
                {
                    ExtractDxfVectorSigned(dxfPath, out groupDxfDx, out groupDxfDy);
                    groupRawPoints = ParseDxfPoints(dxfPath);
                }
                catch (Exception ex)
                {
                    UI?.Log($"  [{cnx}] DXF read error: {ex.Message}. Skipped.");
                    skipped += pairIndices.Count;
                    continue;
                }

                if (groupRawPoints == null || groupRawPoints.Count < 2)
                {
                    UI?.Log($"  [{cnx}] Error: DXF has insufficient points.");
                    skipped += pairIndices.Count;
                    continue;
                }

                foreach (var pair in pairIndices)
                {
                    var ptA = pts[pair.iA];
                    var ptB = pts[pair.iB];

                    try
                    {
                        double ftToMm      = UnitUtils.ConvertFromInternalUnits(1.0, UnitTypeId.Millimeters);
                        double rvtDx       = (ptB.pos.X - ptA.pos.X) * ftToMm;
                        double rvtDy       = (ptB.pos.Y - ptA.pos.Y) * ftToMm;
                        double rvtVano     = Math.Sqrt(rvtDx * rvtDx + rvtDy * rvtDy);
                        double rvtDesnivel = Math.Abs((ptB.pos.Z - ptA.pos.Z) * ftToMm);

                        UI?.Log($"  [{cnx}] RVT → vano: {rvtVano:F1} mm  |  desnivel: {rvtDesnivel:F1} mm");
                        UI?.Log($"  [{cnx}] DXF → dx: {groupDxfDx:F1} mm  dy: {groupDxfDy:F1} mm");

                        bool swapPoints = DetermineSwap(groupDxfDx, groupDxfDy, rvtDx, rvtDy);

                        var ptInicial = swapPoints ? ptB : ptA;
                        var ptFinal   = swapPoints ? ptA : ptB;

                        XYZ pointA = ptInicial.pos;
                        XYZ pointB = ptFinal.pos;

                        List<XYZ> trajectory = CalculateTrajectory(groupRawPoints, pointA, pointB);
                        trajectory = ResampleTrajectory(trajectory, 0.5);

                        double minClearance = 0.5;
                        while (trajectory.Count > 0 && trajectory.First().DistanceTo(pointA) <= minClearance)
                            trajectory.RemoveAt(0);
                        while (trajectory.Count > 0 && trajectory.Last().DistanceTo(pointB) <= minClearance)
                            trajectory.RemoveAt(trajectory.Count - 1);

                        trajectory.Insert(0, pointA);
                        trajectory.Add(pointB);

                        XYZ startTangent = (trajectory[1] - trajectory[0]).Normalize();
                        XYZ endTangent   = (trajectory[trajectory.Count - 1] - trajectory[trajectory.Count - 2]).Normalize();

                        using (Transaction trans = new Transaction(doc, "HMV - Create FlexPipe " + cnx))
                        {
                            trans.Start();

                            FailureHandlingOptions fho = trans.GetFailureHandlingOptions();
                            fho.SetFailuresPreprocessor(new WarningSuppressor());
                            trans.SetFailureHandlingOptions(fho);

                            FlexPipe flexPipe = FlexPipe.Create(
                                doc,
                                systemTypeId,
                                flexType.Id,
                                defaultLevel.Id,
                                startTangent,
                                endTangent,
                                trajectory);

                            if (flexPipe == null)
                            {
                                trans.RollBack();
                                UI?.Log($"  [{cnx}] Error: FlexPipe.Create returned null.");
                                skipped++;
                                continue;
                            }

                            TrySetStringParam(flexPipe, Config.FlexCnxNumberParam,    cnx);
                            TrySetStringParam(flexPipe, Config.FlexEquipoInicialParam, StripEquipmentId(ptInicial.equipoId));
                            TrySetStringParam(flexPipe, Config.FlexEquipoFinalParam,   StripEquipmentId(ptFinal.equipoId));

                            SetConnectedFlag(ptInicial.fi, Config.ConnectedParam);
                            SetConnectedFlag(ptFinal.fi,   Config.ConnectedParam);

                            trans.Commit();
                        }

                        UI?.Log($"  [{cnx}] ✔ Created — Inicial: {ptInicial.equipoId} → Final: {ptFinal.equipoId}" +
                                (swapPoints ? " (swapped)" : ""));
                        created++;
                    }
                    catch (Exception ex)
                    {
                        UI?.Log($"  [{cnx}] Exception: {ex.Message}");
                        skipped++;
                    }
                }
            }

            UI?.Log($"── Done: {created} created, {skipped} skipped ─────────────");
            UI?.SetStatus($"✔ {created} FlexPipe(s) created, {skipped} skipped.");
        }

        // ═══════════════════════════════════════════════════════════════════
        //  DXF HELPERS
        // ═══════════════════════════════════════════════════════════════════

        /// <summary>
        /// Extracts the signed DXF vector (no Math.Abs) from the first valid entity on layer "0".
        /// dxfDx = X(end) - X(start), dxfDy = Y(end) - Y(start) — both in raw DXF units (mm).
        /// </summary>
        private static void ExtractDxfVectorSigned(string filepath, out double dxfDx, out double dxfDy)
        {
            dxfDx = 0;
            dxfDy = 0;

            DxfFile dxf = DxfFile.Load(filepath);
            var entities = dxf.Entities.Where(e => e.Layer == "0").ToList();

            foreach (var ent in entities)
            {
                if (ent is DxfPolyline poly && poly.Vertices.Count > 1)
                {
                    dxfDx = poly.Vertices.Last().Location.X - poly.Vertices.First().Location.X;
                    dxfDy = poly.Vertices.Last().Location.Y - poly.Vertices.First().Location.Y;
                    return;
                }
                if (ent is DxfLwPolyline lwPoly && lwPoly.Vertices.Count > 1)
                {
                    dxfDx = lwPoly.Vertices.Last().X - lwPoly.Vertices.First().X;
                    dxfDy = lwPoly.Vertices.Last().Y - lwPoly.Vertices.First().Y;
                    return;
                }
                if (ent is DxfLine line)
                {
                    dxfDx = line.P2.X - line.P1.X;
                    dxfDy = line.P2.Y - line.P1.Y;
                    return;
                }
            }
        }

        /// <summary>
        /// Decides whether the Inicial/Final assignment should be swapped.
        ///
        /// Strategy:
        ///   Primary axis = Vano (X). Vano must be positive in the Revit model.
        ///   If the Revit X component is negative, swap so it becomes positive.
        ///   If Revit X ≈ 0 (movement is mostly along Y / Desnivel axis), use the Y sign.
        ///   The DXF vector direction confirms the intended orientation.
        ///   When the dot product between DXF and the current Revit vector is negative the
        ///   vectors point in opposite directions → swap.
        /// </summary>
        private static bool DetermineSwap(double dxfDx, double dxfDy, double rvtDx, double rvtDy)
        {
            const double threshold = 1e-3;

            // If DXF has no meaningful vector, fall back to ensuring Revit vano is positive
            bool dxfHasX = Math.Abs(dxfDx) > threshold;
            bool dxfHasY = Math.Abs(dxfDy) > threshold;

            if (!dxfHasX && !dxfHasY)
            {
                // No DXF reference — just make vano (X) positive in Revit
                return rvtDx < -threshold;
            }

            // Primary: use X if DXF has X component
            if (dxfHasX)
            {
                double dot = dxfDx * rvtDx + dxfDy * rvtDy;
                return dot < 0;
            }

            // Secondary: movement is purely on Y axis — use Y sign
            return dxfDy * rvtDy < 0;
        }

        // ── Full DXF point parsing — mirrors ElectricalConnectionCommand.ParseDxfPoints ──

        private static List<XYZ> ParseDxfPoints(string filePath)
        {
            DxfFile dxf = DxfFile.Load(filePath);
            var points = new List<XYZ>();

            foreach (DxfEntity entity in dxf.Entities)
            {
                switch (entity)
                {
                    case DxfLwPolyline lw:
                        foreach (DxfLwPolylineVertex v in lw.Vertices)
                            points.Add(MmToFeet(v.X, v.Y, 0));
                        break;

                    case DxfPolyline poly:
                        foreach (DxfVertex v in poly.Vertices)
                            points.Add(MmToFeet(v.Location.X, v.Location.Y, v.Location.Z));
                        break;

                    case DxfLine line:
                        if (points.Count == 0)
                            points.Add(MmToFeet(line.P1.X, line.P1.Y, line.P1.Z));
                        points.Add(MmToFeet(line.P2.X, line.P2.Y, line.P2.Z));
                        break;

                    case DxfSpline spline:
                        foreach (DxfControlPoint cp in spline.ControlPoints)
                            points.Add(MmToFeet(cp.Point.X, cp.Point.Y, cp.Point.Z));
                        break;
                }
            }

            if (points.Count < 2) return null;

            var cleaned = new List<XYZ> { points[0] };
            for (int i = 1; i < points.Count; i++)
                if (points[i].DistanceTo(cleaned[cleaned.Count - 1]) > 0.001)
                    cleaned.Add(points[i]);

            return cleaned;
        }

        private static XYZ MmToFeet(double x, double y, double z)
        {
            return new XYZ(
                UnitUtils.ConvertToInternalUnits(x, UnitTypeId.Millimeters),
                UnitUtils.ConvertToInternalUnits(y, UnitTypeId.Millimeters),
                UnitUtils.ConvertToInternalUnits(z, UnitTypeId.Millimeters));
        }

        // ── Trajectory — mirrors ElectricalConnectionCommand.CalculateTrajectory ──

        private static List<XYZ> CalculateTrajectory(List<XYZ> dxfPoints, XYZ originA, XYZ originB)
        {
            XYZ dxfStart = dxfPoints[0];
            XYZ dxfEnd   = dxfPoints[dxfPoints.Count - 1];
            XYZ dxfVec   = dxfEnd - dxfStart;
            double dxfLen = dxfVec.GetLength();

            if (dxfLen < 1e-9) throw new InvalidOperationException("DXF start and end are identical.");

            XYZ uDxf = dxfVec.Normalize();
            XYZ vDxf = XYZ.BasisZ.CrossProduct(uDxf).Normalize();

            XYZ revVec = originB - originA;
            double revLen = revVec.GetLength();
            if (revLen < 1e-9) throw new InvalidOperationException("Points A and B are at the same position.");

            XYZ uRev = revVec.Normalize();
            XYZ horizontalNormal = uRev.CrossProduct(XYZ.BasisZ);
            XYZ vRev = horizontalNormal.GetLength() < 1e-9
                ? XYZ.BasisX
                : horizontalNormal.Normalize().CrossProduct(uRev).Normalize();

            double scale = revLen / dxfLen;

            var transformed = new List<XYZ>();
            foreach (XYZ pt in dxfPoints)
            {
                XYZ local    = pt - dxfStart;
                double along   = local.DotProduct(uDxf);
                double deviate = local.DotProduct(vDxf);
                transformed.Add(originA + (uRev * along * scale) + (vRev * deviate * scale));
            }

            transformed[transformed.Count - 1] = originB;
            return transformed;
        }

        private static List<XYZ> ResampleTrajectory(List<XYZ> points, double interval = 0.3)
        {
            double totalLen = 0;
            for (int i = 1; i < points.Count; i++)
                totalLen += points[i].DistanceTo(points[i - 1]);

            if (totalLen < interval * 2) return points;

            var result = new List<XYZ> { points[0] };
            double sinceLastSample = 0;

            for (int i = 1; i < points.Count; i++)
            {
                double segLen = points[i].DistanceTo(points[i - 1]);
                XYZ dir = (points[i] - points[i - 1]).Normalize();
                double consumed = 0;

                while (consumed < segLen)
                {
                    double remaining = interval - sinceLastSample;
                    double available = segLen - consumed;

                    if (available >= remaining)
                    {
                        result.Add(points[i - 1] + dir * (consumed + remaining));
                        consumed += remaining;
                        sinceLastSample = 0;
                    }
                    else
                    {
                        sinceLastSample += available;
                        consumed = segLen;
                    }
                }
            }

            if (result[result.Count - 1].DistanceTo(points[points.Count - 1]) > 0.001)
                result.Add(points[points.Count - 1]);

            return result;
        }

        // ═══════════════════════════════════════════════════════════════════
        //  TYPE RESOLVERS
        // ═══════════════════════════════════════════════════════════════════

        private static FlexPipeType FindFlexPipeType(Document doc, string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return null;

            string[] parts = key.Split(new[] { " : " }, StringSplitOptions.None);
            return parts.Length == 2
                ? new FilteredElementCollector(doc).OfClass(typeof(FlexPipeType))
                    .Cast<FlexPipeType>()
                    .FirstOrDefault(ft => ft.FamilyName == parts[0] && ft.Name == parts[1])
                : new FilteredElementCollector(doc).OfClass(typeof(FlexPipeType))
                    .Cast<FlexPipeType>()
                    .FirstOrDefault(ft => ft.Name == key);
        }

        private static ElementId FindElectricalSystemTypeId(Document doc, string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return null;

            var st = new FilteredElementCollector(doc)
                .OfClass(typeof(PipingSystemType))
                .Cast<PipingSystemType>()
                .FirstOrDefault(s => s.Name.Equals(key, StringComparison.OrdinalIgnoreCase));

            return st?.Id;
        }

        // ═══════════════════════════════════════════════════════════════════
        //  SMALL UTILITIES
        // ═══════════════════════════════════════════════════════════════════

        /// <summary>
        /// Returns greedy minimum-distance index pairs from a list of candidate points.
        /// Pairs where both points share the same equipment ID are excluded.
        /// </summary>
        private static List<(int iA, int iB)> FindOptimalPairIndices(
            List<(ElementId id, FamilyInstance fi, string cnx, string equipoId, XYZ pos)> pts)
        {
            var result = new List<(int, int)>();
            var used   = new HashSet<int>();

            var candidates = new List<(int iA, int iB, double dist)>();
            for (int i = 0; i < pts.Count; i++)
            {
                for (int j = i + 1; j < pts.Count; j++)
                {
                    string idA = ExtractEquipmentId(pts[i].equipoId);
                    string idB = ExtractEquipmentId(pts[j].equipoId);
                    if (!string.IsNullOrWhiteSpace(idA) &&
                        idA.Equals(idB, StringComparison.OrdinalIgnoreCase))
                        continue;
                    candidates.Add((i, j, pts[i].pos.DistanceTo(pts[j].pos)));
                }
            }

            candidates.Sort((a, b) => a.dist.CompareTo(b.dist));

            foreach (var c in candidates)
            {
                if (used.Contains(c.iA) || used.Contains(c.iB)) continue;
                result.Add((c.iA, c.iB));
                used.Add(c.iA);
                used.Add(c.iB);
            }

            return result;
        }

        /// <summary>
        /// Extracts the equipment ElementId portion from a value formatted as "EQUIPO_ID".
        /// "TC_2412545" → "2412545".  If no '_' is found, returns the full string.
        /// </summary>
        private static string ExtractEquipmentId(string equipoValue)
        {
            if (string.IsNullOrWhiteSpace(equipoValue)) return string.Empty;
            int idx = equipoValue.LastIndexOf('_');
            return idx >= 0 ? equipoValue.Substring(idx + 1) : equipoValue;
        }

        /// <summary>
        /// Returns only the equipment name portion, stripping the "_ElementId" suffix.
        /// "TC_2412545" → "TC".  If no '_' is found, returns the full string.
        /// </summary>
        private static string StripEquipmentId(string equipoValue)
        {
            if (string.IsNullOrWhiteSpace(equipoValue)) return string.Empty;
            int idx = equipoValue.LastIndexOf('_');
            return idx >= 0 ? equipoValue.Substring(0, idx) : equipoValue;
        }

        private static Dictionary<string, string> BuildDxfMap(string folder)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string path in Directory.GetFiles(folder, "*.dxf", SearchOption.AllDirectories))
            {
                string name = Path.GetFileNameWithoutExtension(path);
                int idx = name.IndexOf('_');
                string key = idx > 0 ? name.Substring(0, idx) : name;
                if (!map.ContainsKey(key)) map[key] = path;
            }
            return map;
        }

        private static XYZ GetAdaptivePosition(Document doc, FamilyInstance fi)
        {
            try
            {
                IList<ElementId> ptIds =
                    AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(fi);
                if (ptIds != null && ptIds.Count > 0)
                {
                    ReferencePoint rp = doc.GetElement(ptIds[0]) as ReferencePoint;
                    if (rp != null) return rp.Position;
                }
            }
            catch { }
            return (fi.Location as LocationPoint)?.Point ?? XYZ.Zero;
        }

        private static string GetStringParam(Element elem, string paramName)
        {
            if (elem == null || string.IsNullOrWhiteSpace(paramName)) return string.Empty;
            Parameter p = elem.LookupParameter(paramName);
            if (p == null) return string.Empty;
            return p.StorageType == StorageType.String ? p.AsString() ?? string.Empty : string.Empty;
        }

        private static void TrySetStringParam(Element elem, string paramName, string value)
        {
            if (elem == null || string.IsNullOrWhiteSpace(paramName)) return;
            Parameter p = elem.LookupParameter(paramName);
            if (p == null || p.IsReadOnly || p.StorageType != StorageType.String) return;
            p.Set(value ?? string.Empty);
        }

        private static void SetConnectedFlag(FamilyInstance fi, string paramName)
        {
            if (fi == null || string.IsNullOrWhiteSpace(paramName)) return;
            Parameter p = fi.LookupParameter(paramName);
            if (p == null || p.IsReadOnly || p.StorageType != StorageType.Integer) return;
            p.Set(1);
        }
    }
}
