using ACadSharp.Entities;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using HMVTools;
using IxMilia.Dxf;
using IxMilia.Dxf.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Media.Media3D;
using static Autodesk.Revit.DB.SpecTypeId;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class ElectricalConnectionCommand : IExternalCommand
    {
        private static ElectricalConnectionHandler _handler = null;
        private static ExternalEvent _exEvent = null;
        private static ConexionFlexPipeWindow _window = null;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (_handler == null)
            {
                _handler = new ElectricalConnectionHandler();
                _exEvent = ExternalEvent.Create(_handler);
            }

            if (_window != null && _window.IsLoaded)
            {
                _window.Focus();
                return Result.Succeeded;
            }

            _window = new ConexionFlexPipeWindow(commandData.Application, _handler, _exEvent);

            var helper = new System.Windows.Interop.WindowInteropHelper(_window);
            helper.Owner = commandData.Application.MainWindowHandle;

            _window.Show();

            return Result.Succeeded;
        }

        public static void ClearWindow()
        {
            _window = null;
        }
    }

    /// <summary>
    /// IExternalEventHandler that creates a FlexPipe between two points
    /// using a trajectory extracted from a DXF file.
    /// </summary>
    public class ElectricalConnectionHandler : IExternalEventHandler
    {
        // ═══════════════════════════════════════════════════════════════════
        //  PUBLIC PROPERTIES — set by the UI before raising the ExternalEvent
        // ═══════════════════════════════════════════════════════════════════

        public ConexionFlexPipeWindow UI { get; set; }
        public string DxfFilePath { get; set; }
        public ElementId ElementAId { get; set; }
        public ElementId ElementBId { get; set; }
        public string SharedParameterName { get; set; }
        public string SelectedFlexPipeTypeKey { get; set; }
        public string SelectedElectricalSystemTypeKey { get; set; }
        public XYZ PointA { get; set; }
        public XYZ PointB { get; set; }

        public string GetName() => "ElectricalConnectionCommand";

        // ═══════════════════════════════════════════════════════════════════
        //  EXECUTE
        // ═══════════════════════════════════════════════════════════════════
        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;

            if (PointA == null || PointB == null || string.IsNullOrEmpty(DxfFilePath))
            {
                UI?.SetStatus("Error: Missing data to execute.");
                return;
            }

            try
            {
                // ── 1. Validate ──
                UI?.SetStatus("Validating inputs...");
                ValidateInputs(doc);

                // ── 2. Level ──
                Level defaultLevel = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .OrderBy(l => l.Elevation)
                    .FirstOrDefault();

                if (defaultLevel == null)
                    throw new InvalidOperationException("No Level was found in the project.");

                ElementId levelId = defaultLevel.Id;

                // ── 3. Parse DXF ──
                UI?.SetStatus("Parsing DXF geometry...");
                List<XYZ> rawPoints = ParseDxfPoints(DxfFilePath);

                if (rawPoints == null || rawPoints.Count < 2)
                {
                    UI?.SetStatus("Error: DXF is empty or invalid.");
                    return;
                }

                // ── 4. Calculate Trajectory ──
                UI?.SetStatus("Calculating spatial trajectory...");
                List<XYZ> trajectory = CalculateTrajectory(rawPoints, PointA, PointB);
                trajectory = ResampleTrajectory(trajectory, 0.5);

                double minClearance = 0.5;
                while (trajectory.Count > 0 && trajectory.First().DistanceTo(PointA) <= minClearance)
                    trajectory.RemoveAt(0);

                while (trajectory.Count > 0 && trajectory.Last().DistanceTo(PointB) <= minClearance)
                    trajectory.RemoveAt(trajectory.Count - 1);

                trajectory.Insert(0, PointA);
                trajectory.Add(PointB);

                // ── 5. Tangents ──
                XYZ startTangent = (trajectory[1] - trajectory[0]).Normalize();
                XYZ endTangent   = (trajectory[trajectory.Count - 1] - trajectory[trajectory.Count - 2]).Normalize();

                // ── 6. Resolve Types ──
                UI?.SetStatus("Resolving system types...");
                FlexPipeType flexType    = FindFlexPipeType(doc, SelectedFlexPipeTypeKey);
                ElementId flexPipeTypeId = flexType?.Id;
                ElementId systemTypeId   = FindElectricalSystemTypeId(doc, SelectedElectricalSystemTypeKey);

                if (flexPipeTypeId == null || systemTypeId == null)
                {
                    UI?.SetStatus("Error: Missing System Type or FlexPipe Type in the project.");
                    return;
                }

                // ── 7. Create FlexPipe ──
                UI?.SetStatus("Generating FlexPipe...");
                using (Transaction trans = new Transaction(doc, "Create Visual FlexPipe"))
                {
                    trans.Start();

                    FailureHandlingOptions fho = trans.GetFailureHandlingOptions();
                    fho.SetFailuresPreprocessor(new WarningSuppressor());
                    trans.SetFailureHandlingOptions(fho);

                    FlexPipe flexPipe = FlexPipe.Create(
                        doc,
                        systemTypeId,
                        flexPipeTypeId,
                        levelId,
                        startTangent,
                        endTangent,
                        trajectory
                    );

                    if (flexPipe != null && !string.IsNullOrEmpty(SharedParameterName))
                    {
                        string dxfFileName = Path.GetFileName(DxfFilePath);
                        SetParameterValue(flexPipe, SharedParameterName, dxfFileName);
                    }

                    trans.Commit();
                }

                UI?.SetStatus($"✔ Success! FlexPipe created from {Path.GetFileName(DxfFilePath)}");
            }
            catch (Exception ex)
            {
                UI?.SetStatus("Error: " + ex.Message);
            }
        }

        // ═══════════════════════════════════════════════════════════════════
        //  VALIDATION
        // ═══════════════════════════════════════════════════════════════════
        private void ValidateInputs(Document doc)
        {
            if (string.IsNullOrWhiteSpace(DxfFilePath) || !File.Exists(DxfFilePath))
                throw new InvalidOperationException(
                    $"The DXF file doesn't exist or was not specified: '{DxfFilePath}'");

            if (PointA == null || PointB == null)
                throw new InvalidOperationException(
                    "Both points (A and B) must be selected.");

            if (PointA.DistanceTo(PointB) < 1e-6)
                throw new InvalidOperationException(
                    "Point A and Point B cannot be the same location.");

            if (string.IsNullOrWhiteSpace(SelectedFlexPipeTypeKey))
                throw new InvalidOperationException("Must select a Flex Pipe type.");

            if (string.IsNullOrWhiteSpace(SelectedElectricalSystemTypeKey))
                throw new InvalidOperationException("Must select an Electrical System type.");
        }

        // ═══════════════════════════════════════════════════════════════════
        //  CONNECTOR RETRIEVAL
        // ═══════════════════════════════════════════════════════════════════
        private static Connector GetElectricalConnector(FamilyInstance equipment, string label)
        {
            ConnectorManager mgr = equipment?.MEPModel?.ConnectorManager;
            if (mgr == null)
                throw new InvalidOperationException(
                    $"{label} (Id {equipment?.Id}) does not have an MEP Connector Manager.");

            Connector unconnected = null;
            Connector fallback = null;

            foreach (Connector c in mgr.Connectors)
            {
                if (c.Domain != Domain.DomainElectrical) continue;
                if (fallback == null) fallback = c;
                if (!c.IsConnected) { unconnected = c; break; }
            }

            Connector result = unconnected ?? fallback;

            if (result == null)
                throw new InvalidOperationException(
                    $"{label} (Id {equipment.Id}) doesn't have an Electrical Connector.");

            return result;
        }

        // ═══════════════════════════════════════════════════════════════════
        //  DXF PARSING
        // ═══════════════════════════════════════════════════════════════════
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

            if (points.Count < 2)
                throw new InvalidOperationException(
                    $"The DXF file doesn't have enough points (found {points.Count}). Minimum is 2.");

            // Remove duplicate/near-duplicate points
            var cleaned = new List<XYZ> { points[0] };
            for (int i = 1; i < points.Count; i++)
            {
                if (points[i].DistanceTo(cleaned[cleaned.Count - 1]) > 0.001)
                    cleaned.Add(points[i]);
            }

            return cleaned;
        }

        private static XYZ MmToFeet(double xMm, double yMm, double zMm)
        {
            return new XYZ(
                UnitUtils.ConvertToInternalUnits(xMm, UnitTypeId.Millimeters),
                UnitUtils.ConvertToInternalUnits(yMm, UnitTypeId.Millimeters),
                UnitUtils.ConvertToInternalUnits(zMm, UnitTypeId.Millimeters)
            );
        }

        // ═══════════════════════════════════════════════════════════════════
        //  TRAJECTORY CALCULATION
        // ═══════════════════════════════════════════════════════════════════
        private static List<XYZ> CalculateTrajectory(List<XYZ> dxfPoints, XYZ originA, XYZ originB)
        {
            XYZ dxfStart = dxfPoints[0];
            XYZ dxfEnd   = dxfPoints[dxfPoints.Count - 1];
            XYZ dxfVec   = dxfEnd - dxfStart;
            double dxfLen = dxfVec.GetLength();

            if (dxfLen < 1e-9)
                throw new InvalidOperationException("DXF start and end points are identical.");

            XYZ uDxf = dxfVec.Normalize();
            XYZ vDxf = XYZ.BasisZ.CrossProduct(uDxf).Normalize();

            XYZ revVec = originB - originA;
            double revLen = revVec.GetLength();

            if (revLen < 1e-9)
                throw new InvalidOperationException("Points A and B are in the exact same position.");

            XYZ uRev = revVec.Normalize();

            XYZ horizontalNormal = uRev.CrossProduct(XYZ.BasisZ);
            XYZ vRev = horizontalNormal.GetLength() < 1e-9
                ? XYZ.BasisX
                : horizontalNormal.Normalize().CrossProduct(uRev).Normalize();

            double scale = revLen / dxfLen;

            var transformed = new List<XYZ>();
            foreach (XYZ pt in dxfPoints)
            {
                XYZ local = pt - dxfStart;
                double along    = local.DotProduct(uDxf);
                double deviate  = local.DotProduct(vDxf);
                transformed.Add(originA + (uRev * along * scale) + (vRev * deviate * scale));
            }

            transformed[transformed.Count - 1] = originB;
            return transformed;
        }

        // ═══════════════════════════════════════════════════════════════════
        //  FLEXPIPE TYPE HELPERS
        // ═══════════════════════════════════════════════════════════════════
        private static FlexPipeType FindFlexPipeType(Document doc, string key)
        {
            if (string.IsNullOrWhiteSpace(key))
                throw new InvalidOperationException("Flex pipe type not specified.");

            string[] parts = key.Split(new[] { " : " }, StringSplitOptions.None);

            FlexPipeType result = parts.Length == 2
                ? new FilteredElementCollector(doc).OfClass(typeof(FlexPipeType))
                    .Cast<FlexPipeType>()
                    .FirstOrDefault(ft => ft.FamilyName == parts[0] && ft.Name == parts[1])
                : new FilteredElementCollector(doc).OfClass(typeof(FlexPipeType))
                    .Cast<FlexPipeType>()
                    .FirstOrDefault(ft => ft.Name == key);

            if (result == null)
                throw new InvalidOperationException($"FlexPipeType '{key}' not found in model.");

            return result;
        }

        private static ElementId FindElectricalSystemTypeId(Document doc, string key)
        {
            if (string.IsNullOrWhiteSpace(key))
                throw new InvalidOperationException("Electrical system type not specified.");

            var systemType = new FilteredElementCollector(doc)
                .OfClass(typeof(PipingSystemType))
                .Cast<PipingSystemType>()
                .FirstOrDefault(st => st.Name.Equals(key, StringComparison.OrdinalIgnoreCase));

            if (systemType == null)
                throw new InvalidOperationException(
                    $"PipingSystemType '{key}' not found in the model.");

            return systemType.Id;
        }

        // ═══════════════════════════════════════════════════════════════════
        //  PARAMETER SETTER
        // ═══════════════════════════════════════════════════════════════════
        private static void SetParameterValue(Element element, string paramName, string value)
        {
            if (string.IsNullOrWhiteSpace(paramName) || string.IsNullOrWhiteSpace(value))
                return;

            Parameter p = element.LookupParameter(paramName);

            if (p == null)
                throw new InvalidOperationException(
                    $"Parameter '{paramName}' does not exist on FlexPipe (Id {element.Id}).");

            if (p.IsReadOnly)
                throw new InvalidOperationException($"Parameter '{paramName}' is read-only.");

            switch (p.StorageType)
            {
                case StorageType.String:
                    p.Set(value);
                    break;
                case StorageType.Double:
                    if (double.TryParse(value, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out double dVal))
                        p.Set(dVal);
                    else
                        throw new InvalidOperationException(
                            $"Cannot convert '{value}' to double for parameter '{paramName}'.");
                    break;
                case StorageType.Integer:
                    if (int.TryParse(value, out int iVal))
                        p.Set(iVal);
                    else
                        throw new InvalidOperationException(
                            $"Cannot convert '{value}' to integer for parameter '{paramName}'.");
                    break;
                default:
                    throw new InvalidOperationException(
                        $"Unsupported storage type ({p.StorageType}) for parameter '{paramName}'.");
            }
        }

        // ═══════════════════════════════════════════════════════════════════
        //  STUB: CLOUD UPDATE PLACEHOLDER
        // ═══════════════════════════════════════════════════════════════════
        private void HandleExistingPipeUpdate(Document doc, FamilyInstance equipA, FamilyInstance equipB)
        {
            // TODO: Cloud-based update logic
        }

        // ═══════════════════════════════════════════════════════════════════
        //  RESAMPLE (utility, currently unused)
        // ═══════════════════════════════════════════════════════════════════
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
                    double remaining  = interval - sinceLastSample;
                    double available  = segLen - consumed;

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
    }
}