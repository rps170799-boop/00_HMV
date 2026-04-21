using ACadSharp.Entities;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using HMVTools;
using IxMilia.Dxf;
using IxMilia.Dxf.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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

            // Si la ventana ya existe, la traemos al frente
            if (_window != null && _window.IsLoaded)
            {
                _window.Focus();
                return Result.Succeeded;
            }

            _window = new ConexionFlexPipeWindow(commandData.Application, _handler, _exEvent);

            // Vincular con la ventana principal de Revit para que no se pierda detrás
            var helper = new System.Windows.Interop.WindowInteropHelper(_window);
            helper.Owner = commandData.Application.MainWindowHandle;

            _window.Show(); // Modeless: permite interactuar con Revit

            return Result.Succeeded;
        }

        // Método para limpiar la referencia cuando el usuario cierra la ventana
        public static void ClearWindow()
        {
            _window = null;
        }

    }
        /// <summary>
        /// IExternalEventHandler that creates a FlexPipe between two Electrical Equipment
        /// instances using a trajectory extracted from a DXF file.
        /// 
        /// ── Workflow ──
        /// 1. Parse DXF → ordered points (mm → feet).
        /// 2. Transform points: anchor at Connector A, orient toward Connector B.
        /// 3. Snap last point to Connector B exactly.
        /// 4. Create FlexPipe along the trajectory.
        /// 5. Connect MEP connectors on both ends.
        /// 6. Write DXF filename into a Shared Parameter.
        /// </summary>
        public class ElectricalConnectionHandler : IExternalEventHandler
        {
        // ═══════════════════════════════════════════════════════════════════
        //  PUBLIC PROPERTIES — set by the UI before raising the ExternalEvent
        // ═══════════════════════════════════════════════════════════════════

        /// <summary>Reference to the modeless window for status updates.</summary>
             public ConexionFlexPipeWindow UI { get; set; }  // Replace 'object' with your Vista type

            /// <summary>Full path to the .dxf file containing the trajectory.</summary>
            public string DxfFilePath { get; set; }

            /// <summary>ElementId of Electrical Equipment A (start).</summary>
            public ElementId ElementAId { get; set; }

            /// <summary>ElementId of Electrical Equipment B (end).</summary>
            public ElementId ElementBId { get; set; }

            /// <summary>Name of the Shared Parameter to store the DXF filename.</summary>
            public string SharedParameterName { get; set; }

            /// <summary>"Family : Type" key for the FlexPipeType to use.</summary>
            public string SelectedFlexPipeTypeKey { get; set; }

            /// <summary>"Family : Type" key or name for the ElectricalSystemType.</summary>
            public string SelectedElectricalSystemTypeKey { get; set; }

            public string GetName() => "ElectricalConnectionCommand";

        // ═══════════════════════════════════════════════════════════════════
        //  EXECUTE — high-level recipe
        // ═══════════════════════════════════════════════════════════════════
        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;

            // Safety validation in case the event is fired with empty data
            if (ElementAId == null || ElementBId == null || string.IsNullOrEmpty(DxfFilePath))
            {
                UI?.SetStatus("Error: Missing data to execute.");
                return;
            }

            try
            {
                // ── 1. Validate inputs ──────────────────────────────────
                UI?.SetStatus("Validating inputs...");
                ValidateInputs(doc);

                // ── 2. Retrieve electrical connectors from both elements ─
                UI?.SetStatus("Retrieving electrical connectors...");
                FamilyInstance equipA = doc.GetElement(ElementAId) as FamilyInstance;
                FamilyInstance equipB = doc.GetElement(ElementBId) as FamilyInstance;

                Connector connectorA = GetElectricalConnector(equipA, "Element A");
                Connector connectorB = GetElectricalConnector(equipB, "Element B");

                // Level resolution fallback
                Element elemA = doc.GetElement(ElementAId);
                ElementId validLevelId = elemA.LevelId;

                if (validLevelId == ElementId.InvalidElementId)
                {
                    Level fallbackLevel = new FilteredElementCollector(doc)
                        .OfClass(typeof(Level))
                        .Cast<Level>()
                        .OrderBy(l => l.Elevation) // Take the lowest level
                        .FirstOrDefault();

                    if (fallbackLevel == null)
                    {
                        throw new InvalidOperationException("No Level was found in the project.");
                    }

                    validLevelId = fallbackLevel.Id;
                }

                // ── 3. Parse DXF → ordered points (mm → decimal feet) ──
                UI?.SetStatus("Parsing DXF geometry...");
                List<XYZ> rawPoints = ParseDxfPoints(DxfFilePath);

                // ── 4. Transform trajectory into Revit coordinate space ─
                UI?.SetStatus("Calculating spatial trajectory...");
                List<XYZ> trajectory = CalculateTrajectory(rawPoints, connectorA.Origin, connectorB.Origin);
                trajectory = ResampleTrajectory(trajectory, 0.3);


                // ── 5. Stub: Cloud-based update logic ───────────────────
                UI?.SetStatus("Checking existing pipe connections...");
                HandleExistingPipeUpdate(doc, equipA, equipB);

                // ── 6. Create FlexPipe and connect ──────────────────────
                UI?.SetStatus("Generating 3D FlexPipe...");

                using (Transaction tx = new Transaction(doc, "HMV Tools – FlexPipe DXF"))
                {
                    tx.Start();

                    // Suppress non-critical warnings
                    FailureHandlingOptions fho = tx.GetFailureHandlingOptions();
                    fho.SetFailuresPreprocessor(new WarningSuppressor());
                    tx.SetFailureHandlingOptions(fho);

                    // 6a. Resolve types
                    FlexPipeType flexType = FindFlexPipeType(doc, SelectedFlexPipeTypeKey);
                    ElementId systemTypeId = FindElectricalSystemTypeId(doc, SelectedElectricalSystemTypeKey);

                    XYZ startTangent = (trajectory[1] - trajectory[0]).Normalize();

                    // endTangent: Vector from the penultimate point to the last point (Connector B)
                    XYZ endTangent = (trajectory[trajectory.Count - 1] - trajectory[trajectory.Count - 2]).Normalize();

                    // 6b. Create the FlexPipe
                    FlexPipe pipe = FlexPipe.Create(
                        doc,
                        systemTypeId,
                        flexType.Id,
                        validLevelId,
                        startTangent,
                        endTangent,
                        trajectory);

                    // 6c. Connect MEP endpoints
                    //ConnectFlexPipeToEquipment(pipe, connectorA, connectorB);

                    // 6d. Fill the Shared Parameter with the DXF filename
                    if (!string.IsNullOrEmpty(SharedParameterName))
                    {
                        string dxfFileName = System.IO.Path.GetFileName(DxfFilePath);
                        SetParameterValue(pipe, SharedParameterName, dxfFileName);
                    }

                    tx.Commit();
                }

                // ── 7. Done ─────────────────────────────────────────────
                // Print success directly to the UI instead of blocking Revit with a TaskDialog
                UI?.SetStatus($"✔ Success! FlexPipe created from {System.IO.Path.GetFileName(DxfFilePath)}");
            }
            catch (Exception ex)
            {
                // Print the error gracefully in the UI to allow the user to keep working
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
                        $"The DXF File doesn´t exist or was not specified: '{DxfFilePath}'");

                if (ElementAId == null || ElementBId == null)
                    throw new InvalidOperationException(
                        "Both equipment elements (Element A and Element B) must be selected.");
                if (ElementAId == ElementBId)
                    throw new InvalidOperationException(
                        "Element A and Element B can´t be the same");

                Element a = doc.GetElement(ElementAId);
                Element b = doc.GetElement(ElementBId);

                if (a == null || b == null)
                    throw new InvalidOperationException(
                        "One or both elements doesn´t exist in the model");

                if (string.IsNullOrWhiteSpace(SelectedFlexPipeTypeKey))
                    throw new InvalidOperationException("Must select a Flex Pipe.");

                if (string.IsNullOrWhiteSpace(SelectedElectricalSystemTypeKey))
                    throw new InvalidOperationException("Must selec a electrical System.");
            }

            // ═══════════════════════════════════════════════════════════════════
            //  CONNECTOR RETRIEVAL
            // ═══════════════════════════════════════════════════════════════════

            /// <summary>
            /// Gets the first available (unconnected) electrical connector from
            /// a FamilyInstance.  Throws if none found.
            /// </summary>
            private static Connector GetElectricalConnector(FamilyInstance equipment, string label)
            {
                ConnectorManager mgr = equipment?.MEPModel?.ConnectorManager;
                if (mgr == null)
                    throw new InvalidOperationException(
                        $"{label} (Id {equipment?.Id}) doesnpt have MEP Conncetor Manager. " +
                        "Verify that its a Electrical Equipment with MEP connectors.");

                // Prefer an unconnected electrical connector; fall back to any electrical connector
                Connector unconnected = null;
                Connector fallback = null;

                foreach (Connector c in mgr.Connectors)
                {
                    if (c.Domain != Domain.DomainElectrical) continue;

                    if (fallback == null) fallback = c;

                    if (!c.IsConnected)
                    {
                        unconnected = c;
                        break;
                    }
                }

                Connector result = unconnected ?? fallback;

                if (result == null)
                    throw new InvalidOperationException(
                        $"{label} (Id {equipment.Id}) doesn´t have Electrical Connector. " +
                        "Verify the family.");

                return result;
            }

            // ═══════════════════════════════════════════════════════════════════
            //  DXF PARSING
            // ═══════════════════════════════════════════════════════════════════

            /// <summary>
            /// Reads the DXF file silently using IxMilia.Dxf.
            /// Extracts ordered vertices from LWPOLYLINE, POLYLINE, or LINE entities.
            /// Converts all coordinates from millimetres to Revit internal units (decimal feet).
            /// </summary>
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
                            {
                                points.Add(MmToFeet(v.X, v.Y, 0));
                            }
                            break;

                        case DxfPolyline poly:
                            foreach (DxfVertex v in poly.Vertices)
                            {
                                points.Add(MmToFeet(v.Location.X, v.Location.Y, v.Location.Z));
                            }
                            break;

                        case DxfLine line:
                            // Only add P1 to avoid duplicating shared endpoints;
                            // P2 will be P1 of the next line or handled after the loop.
                            if (points.Count == 0)
                                points.Add(MmToFeet(line.P1.X, line.P1.Y, line.P1.Z));
                            points.Add(MmToFeet(line.P2.X, line.P2.Y, line.P2.Z));
                            break;

                        case DxfSpline spline:
                            // Use control points as an approximation
                            foreach (DxfControlPoint cp in spline.ControlPoints)
                            {
                                points.Add(MmToFeet(cp.Point.X, cp.Point.Y, cp.Point.Z));
                            }
                            break;
                    }
                }

                if (points.Count < 2)
                    throw new InvalidOperationException(
                        $"The DXF File doesn´t have enough points to get a path. " +
                        $"(found {points.Count}). IT MUST BE AT LEAST 2 POINTS.");

            // ── Smart Cleanup: Remove collinear points and prevent micro-segments ──

            // At the end of ParseDxfPoints, REPLACE the entire smart cleanup block with:
            var cleaned = new List<XYZ> { points[0] };
            for (int i = 1; i < points.Count; i++)
            {
                if (points[i].DistanceTo(cleaned.Last()) > 0.001)
                    cleaned.Add(points[i]);
            }
            return cleaned;
        }

        /// <summary>
        /// Converts mm coordinates to Revit internal units (decimal feet).
        /// Uses UnitUtils for Revit 2023 (ForgeTypeId API).
        /// </summary>
        private static XYZ MmToFeet(double xMm, double yMm, double zMm)
            {
                double x = UnitUtils.ConvertToInternalUnits(xMm, UnitTypeId.Millimeters);
                double y = UnitUtils.ConvertToInternalUnits(yMm, UnitTypeId.Millimeters);
                double z = UnitUtils.ConvertToInternalUnits(zMm, UnitTypeId.Millimeters);
                return new XYZ(x, y, z);
            }

        // ═══════════════════════════════════════════════════════════════════
        //  TRAJECTORY CALCULATION & COORDINATE MAPPING (Z-PLANE FORCED)
        // ═══════════════════════════════════════════════════════════════════

        private static List<XYZ> CalculateTrajectory(List<XYZ> dxfPoints, XYZ originA, XYZ originB)
        {
            // ── 1. Analyze DXF natural direction (first → last) ────────────
            XYZ dxfStart = dxfPoints.First();
            XYZ dxfEnd = dxfPoints.Last();
            XYZ dxfVec = dxfEnd - dxfStart;
            double dxfLen = dxfVec.GetLength();

            if (dxfLen < 1e-9)
                throw new InvalidOperationException(
                    "The start and end points of the DXF are the same. " +
                    "Please verify the DXF file.");

            XYZ uDxf = dxfVec.Normalize();

            // Perpendicular vector in the DXF plane (assuming it was drawn flat in CAD)
            XYZ zDxf = XYZ.BasisZ;
            XYZ vDxf = zDxf.CrossProduct(uDxf).Normalize();

            // ── 2. Analyze Revit target direction (A → B) ──────────────────
            XYZ revVec = originB - originA;
            double revLen = revVec.GetLength();

            if (revLen < 1e-9)
                throw new InvalidOperationException(
                    "Connectors A and B are in the exact same position.");

            XYZ uRev = revVec.Normalize(); // The straight axis between connectors

            // ── 3. Force the deviation plane (Z-Axis / Elevation) ──────────
            XYZ vRev;
            XYZ globalZ = XYZ.BasisZ;
            XYZ horizontalNormal = uRev.CrossProduct(globalZ);

            // If the straight line is perfectly vertical, the Z-axis is already taken, 
            // so we force the curve's deviation to the X-axis to avoid calculation errors.
            if (horizontalNormal.GetLength() < 1e-9)
            {
                vRev = XYZ.BasisX;
            }
            else
            {
                // Get a vector perpendicular to the straight line, strictly in the Z-elevation plane
                vRev = horizontalNormal.Normalize().CrossProduct(uRev).Normalize();
            }

            // Calculate proportional scale to fit the DXF exactly between the connectors
            double scale = revLen / dxfLen;

            // ── 4. Build the transformed points ────────────────────────────
            List<XYZ> transformed = new List<XYZ>();

            foreach (XYZ pt in dxfPoints)
            {
                // Shift DXF to its own local origin
                XYZ localPt = pt - dxfStart;

                // Project: How far along the straight line?
                double distanceAlong = localPt.DotProduct(uDxf);

                // Project: How far does it deviate from the line? (The curve's belly)
                double deviation = localPt.DotProduct(vDxf);

                // Rebuild the point in Revit forcing the deviation onto our vRev vector (Vertical)
                XYZ worldPt = originA
                            + (uRev * distanceAlong * scale)
                            + (vRev * deviation * scale);

                transformed.Add(worldPt);
            }

            // ── 5. CRITICAL: Force last point to exact Connector B origin ──
            transformed[transformed.Count - 1] = originB;

            return transformed;
        }


        
        /// <summary>
        /// Resamples the polyline at uniform intervals along its length.
        /// Guarantees consistent segment length and smooth angles.
        /// Interval should be ~5x pipe diameter minimum.
        /// </summary>
        private static List<XYZ> ResampleTrajectory(List<XYZ> points, double interval = 0.3)
        {
            // Calculate total polyline length
            double totalLen = 0;
            for (int i = 1; i < points.Count; i++)
                totalLen += points[i].DistanceTo(points[i - 1]);

            if (totalLen < interval * 2)
                return points; // Too short to resample

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
                        // Place a sample point
                        XYZ sample = points[i - 1] + dir * (consumed + remaining);
                        result.Add(sample);
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

            // Always end exactly at the last point
            if (result.Last().DistanceTo(points.Last()) > 0.001)
                result.Add(points.Last());

            

            return result;
        }

        
            // ═══════════════════════════════════════════════════════════════════
            //  STUB: CLOUD-BASED UPDATE LOGIC
            // ═══════════════════════════════════════════════════════════════════

            /// <summary>
            /// TODO: Cloud-based Update logic to be implemented here.
            /// 
            /// This method will eventually:
            ///   1. Query whether a FlexPipe already exists between equipA and equipB.
            ///   2. If found, delete or deactivate the existing pipe.
            ///   3. Sync the action with the cloud service.
            ///   4. Allow the main workflow to proceed with fresh creation.
            ///   
            /// Currently a no-op placeholder.
            /// </summary>
            private void HandleExistingPipeUpdate(
                Document doc, FamilyInstance equipA, FamilyInstance equipB)
            {
                // TODO: Cloud-based Update logic to be implemented here.
                // No deletion, search, or recreation logic at this stage.
            }

            // ═══════════════════════════════════════════════════════════════════
            //  FLEXPIPE CREATION HELPERS
            // ═══════════════════════════════════════════════════════════════════

            /// <summary>
            /// Finds a FlexPipeType by its "Family : Type" key.
            /// </summary>
            private static FlexPipeType FindFlexPipeType(Document doc, string key)
            {
                if (string.IsNullOrWhiteSpace(key))
                    throw new InvalidOperationException("Flex pipe type doesn´t specified.");

                string[] parts = key.Split(new[] { " : " }, StringSplitOptions.None);

                FlexPipeType result;

                if (parts.Length == 2)
                {
                    result = new FilteredElementCollector(doc)
                        .OfClass(typeof(FlexPipeType))
                        .Cast<FlexPipeType>()
                        .FirstOrDefault(ft =>
                            ft.FamilyName == parts[0] && ft.Name == parts[1]);
                }
                else
                {
                    // Fallback: match by name only
                    result = new FilteredElementCollector(doc)
                        .OfClass(typeof(FlexPipeType))
                        .Cast<FlexPipeType>()
                        .FirstOrDefault(ft => ft.Name == key);
                }

                if (result == null)
                    throw new InvalidOperationException(
                        $"FlexPipeType '{key}' doesn´t found in model.");

                return result;
            }

            /// <summary>
            /// Finds an Electrical System Type by name.
            /// Returns its ElementId for use in FlexPipe.Create().
            /// </summary>
            private static ElementId FindElectricalSystemTypeId(Document doc, string key)
            {
                if (string.IsNullOrWhiteSpace(key))
                    throw new InvalidOperationException("Electrical System doesn´t specified.");

                // PipingSystemType is the base class Revit uses for pipe system classification.
                // For electrical conduit/pipe routing we look for a matching system type by name.
                var systemType = new FilteredElementCollector(doc)
                    .OfClass(typeof(PipingSystemType))
                    .Cast<PipingSystemType>()
                    .FirstOrDefault(st =>
                        st.Name.Equals(key, StringComparison.OrdinalIgnoreCase));

                if (systemType == null)
                    throw new InvalidOperationException(
                        $"PipingSystemType '{key}' doesn´t found in the model. " +
                        "Verify that it exists in Mechanical Settings > System Types.");

                return systemType.Id;
            }

            

            

            // ═══════════════════════════════════════════════════════════════════
            //  PARAMETER SETTER
            // ═══════════════════════════════════════════════════════════════════

            /// <summary>
            /// Sets a Shared/Instance parameter on the FlexPipe element.
            /// Supports String, Double, and Integer storage types.
            /// </summary>
            private static void SetParameterValue(Element element, string paramName, string value)
            {
                if (string.IsNullOrWhiteSpace(paramName) || string.IsNullOrWhiteSpace(value))
                    return;

                Parameter p = element.LookupParameter(paramName);

                if (p == null)
                    throw new InvalidOperationException(
                        $"Shared Parameter '{paramName}' does not exist on the element " +
                        $"FlexPipe (Id {element.Id}). Verify that it is loaded in the project.");

                if (p.IsReadOnly)
                    throw new InvalidOperationException(
                        $"The parameter '{paramName}' is read-only.");

                switch (p.StorageType)
                {
                    case StorageType.String:
                        p.Set(value);
                        break;

                    case StorageType.Double:
                        if (double.TryParse(value,
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture,
                            out double dVal))
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
        }
    
}