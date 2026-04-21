using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration.Assemblies;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace HMVTools
{
    public class App : IExternalApplication
    {
        // ══════════════════════════════════════════════════════════
        //  PROACTIVE TRACKING STATE
        // ══════════════════════════════════════════════════════════

        /// <summary>
        /// Captures the source path during FamilyLoadingIntoDocument
        /// so it can be read in FamilyLoadedIntoDocument (which lacks
        /// a FamilyPath property in Revit 2023).
        /// Key = FamilyName, Value = FamilyPath.
        /// </summary>
        private static readonly ConcurrentDictionary<string, string> _pendingPaths =
            new ConcurrentDictionary<string, string>();

        /// <summary>
        /// Guard against infinite recursion: InjectVisibleParameter
        /// reloads the family, which fires FamilyLoadedIntoDocument
        /// again. Families in this set are currently being stamped.
        /// </summary>
        private static readonly HashSet<string> _stampingInProgress =
            new HashSet<string>();

        /// <summary>ExternalEvent used to defer Revit DB writes out of event handlers.</summary>
        private static ExternalEvent _stampEvent;
        private static StampFamilyHandler _stampHandler;

        /// <summary>
        /// Master switch: set to false to disable proactive tracking
        /// without removing the event wiring.
        /// </summary>
        public static bool ProactiveTrackingEnabled { get; set; } = true;

        /// <summary>
        /// When true, the heavyweight parameter injection (EditFamily +
        /// reload) runs on each tracked load. When false, only the
        /// lightweight ExtensibleStorage stamp is written.
        /// </summary>
        public static bool InjectParameterOnLoad { get; set; } = false;

        /// <summary>Reference to the recursion guard set, used by the handler.</summary>
        internal static HashSet<string> StampingInProgress => _stampingInProgress;

        // ══════════════════════════════════════════════════════════
        //  STARTUP
        // ══════════════════════════════════════════════════════════

        public Result OnStartup(UIControlledApplication app)
        {
            app.CreateRibbonTab("HMV Tools");
            string path = Assembly.GetExecutingAssembly().Location;

            // ── DWG Panel ──
            RibbonPanel panelDwg = app.CreateRibbonPanel("HMV Tools", "DWG");

            PushButtonData btnConvert = new PushButtonData(
                "DwgConvert",
                "DWG\nConvert",
                path,
                "HMVTools.DwgConvertCommand");
            btnConvert.ToolTip = "Convert DWG lines and texts to HMV standards";
            btnConvert.LongDescription =
                "Step 1 – Convert Lines: extracts geometry from the DWG import "
                + "and creates HMV_LINEA detail lines.\n"
                + "Step 2 – Partial Explode the DWG manually in Revit.\n"
                + "Step 3 – Standardize All: re-styles exploded lines to HMV_LINEA, "
                + "converts texts to HMV_General_<size> <font>, and purges "
                + "unused DWG-imported styles from the project.";
            PushButton btn = panelDwg.AddItem(btnConvert) as PushButton;
            BitmapImage icon = LoadImage("HMVTools.Resources.dwg_32.png");
            if (icon != null) btn.LargeImage = icon;

            PushButtonData btnDwg3D = new PushButtonData(
                "Dwg3DToShape",
                "3D DWG\nto Shape",
                path,
                "HMVTools.Dwg3DToShapeCommand");
            btnDwg3D.ToolTip = "Convert imported 3D DWG geometry to a native DirectShape";
            btnDwg3D.LongDescription =
                "Extracts Solids and Meshes from an imported DWG element "
                + "and creates a single DirectShape in the chosen category.\n\n"
                + "Features:\n"
                + "• Recursive geometry extraction (handles nested GeometryInstances)\n"
                + "• Volume threshold to skip bolts and small parts\n"
                + "• TessellatedShapeBuilder for mesh-to-solid conversion\n"
                + "• Detailed report of kept/skipped/converted geometry";
            PushButton btnD3 = panelDwg.AddItem(btnDwg3D) as PushButton;
            BitmapImage iconD3 = LoadImage("HMVTools.Resources.dwg3d_32.png");
            if (iconD3 != null) btnD3.LargeImage = iconD3;

            // ── Family Control Tools Panel ──
            RibbonPanel panelFamily = app.CreateRibbonPanel("HMV Tools", "Family Control Tools");

            PushButtonData btnDuctos = new PushButtonData(
                "DuctosEditor",
                "Ductos\nEditor",
                path,
                "HMVTools.USDTuberiaConfigCommand");
            btnDuctos.ToolTip = "Configurador visual de tuberías USD";
            btnDuctos.LongDescription =
                "Seleccione una o más instancias de la familia USD y ejecute "
                + "este comando para editar visibilidad, diámetro y distancias "
                + "de las 45 tuberías (5×9) desde una interfaz visual.\n\n"
                + "Funciones:\n"
                + "• Grilla 5×9 con checkbox y diámetro por tubería\n"
                + "• Control de distancias X/Y entre ejes\n"
                + "• Herramientas masivas: toggle fila/columna, ø global\n"
                + "• Aplica cambios a todas las instancias seleccionadas.";
            PushButton btnD = panelFamily.AddItem(btnDuctos) as PushButton;
            BitmapImage iconDuctos = LoadImage("HMVTools.Resources.ductos_32.png");
            if (iconDuctos != null) btnD.LargeImage = iconDuctos;

            PushButtonData btnRefreshZ = new PushButtonData(
                "RefreshZToFloor",
                "Refresh Z\nto Floor",
                path,
                "HMVTools.RefreshZToFloorCommand");
            btnRefreshZ.ToolTip = "Snap element Z to the top of a linked floor";
            btnRefreshZ.LongDescription =
                "Adjusts the Z elevation of selected elements (foundations, etc.) "
                + "so they sit exactly on top of a floor from a linked Revit model.\n\n"
                + "Workflow:\n"
                + "1. Move elements to their new X,Y position.\n"
                + "2. Select the moved elements (or pre-select before running).\n"
                + "3. Run this command and pick the Revit link with the floor.\n"
                + "4. The tool shoots a vertical ray at each element's X,Y and "
                + "snaps its Z to the top face of the nearest linked floor.\n\n"
                + "Works with any point-based element (foundations, equipment, etc.).";
            PushButton btnRZ = panelFamily.AddItem(btnRefreshZ) as PushButton;
            BitmapImage iconRZ = LoadImage("HMVTools.Resources.refreshz_32.png");
            if (iconRZ != null) btnRZ.LargeImage = iconRZ;

            PushButtonData btnMultiParam = new PushButtonData(
                "MultiParamEditor",
                "Multi InstParam\nEditor",
                path,
                "HMVTools.MultiParamEditorCommand");
            btnMultiParam.ToolTip = "Edit multiple instance parameters simultaneously";
            btnMultiParam.LongDescription =
                "Select one or more family instances, then edit multiple instance "
                + "parameters in a single transaction to avoid constraint conflicts.\n\n"
                + "Features:\n"
                + "• Dynamic parameter rows with dropdown and value input\n"
                + "• Auto-detects parameter units (mm, m, °, etc.)\n"
                + "• Shows current values in grey (comma-separated if they vary)\n"
                + "• Multiple families: choose Common Parameters or Each Family mode\n"
                + "• Each Family mode persists rows independently per family\n"
                + "• All parameters applied simultaneously in one transaction";
            PushButton btnMP = panelFamily.AddItem(btnMultiParam) as PushButton;
            BitmapImage iconMP = LoadImage("HMVTools.Resources.multiparam_32.png");
            if (iconMP != null) btnMP.LargeImage = iconMP;

           

            PushButtonData btnFoundation = new PushButtonData(
                "FoundationControl",
                "Foundation\nControl",
                path,
                "HMVTools.FoundationLayoutCommand");
            btnFoundation.ToolTip = "Foundation Layout & Elevation Manager";
            btnFoundation.LongDescription =
                "Adjust the vertical position of structural foundations to match "
                + "a desired NTCE (top of concrete elevation) in survey coordinates.\n\n"
                + "Workflow:\n"
                + "1. Select structural foundations in the model.\n"
                + "2. Choose the Revit link containing the reference floors.\n"
                + "3. View the foundations on a scale-to-fit plan canvas.\n"
                + "4. Review the detected NAP (floor level) for each foundation.\n"
                + "5. Edit the desired NTCE elevation (survey meters).\n"
                + "6. Apply to move foundations vertically in one transaction.\n\n"
                + "Features:\n"
                + "• Dense downward raycasting for accurate floor detection on slopes\n"
                + "• Displays NAP in both survey and project coordinates\n"
                + "• Visual plan canvas with interactive selection\n"
                + "• Validates all inputs before applying changes";
            PushButton btnFC = panelFamily.AddItem(btnFoundation) as PushButton;
            BitmapImage iconFC = LoadImage("HMVTools.Resources.foundation_32.png");
            if (iconFC != null) btnFC.LargeImage = iconFC;


            PushButtonData btnSignageHosting = new PushButtonData(
                "SignageHosting",
                "Host\nSignage",
                path,
                "HMVTools.SignageHostingCommand");
            btnSignageHosting.ToolTip = "Batch host signage to nested pedestal components";
            btnSignageHosting.LongDescription =
                "Automatically finds the closest Electrical Equipment to selected pedestals, "
                + "hosts a Signage family on the top face of a nested pedestal component, "
                + "and syncs a chosen parameter value.\n\n"
                + "Workflow:\n"
                + "1. Pre-select Pedestals in the model (or select nothing to process all).\n"
                + "2. Run the tool and select your Signage family.\n"
                + "3. Define the internal pedestal component name and the parameters to sync.";

            PushButton btnSH = panelFamily.AddItem(btnSignageHosting) as PushButton;
            BitmapImage iconSH = LoadImage("HMVTools.Resources.signage_32.png"); 
            if (iconSH != null) btnSH.LargeImage = iconSH;


           
            PushButtonData btnConstraints = new PushButtonData(
                "ConstraintsRelease",
                "Constraints\nRelease",
                path,
                "HMVTools.ConstraintsReleaseCommand");
            btnConstraints.ToolTip = "Release all constraints from selected objects";
            btnConstraints.LongDescription = "Removes alignment and parameter constraints (locked dimensions and labels) from the selected instances.";
            PushButton btnCR = panelFamily.AddItem(btnConstraints) as PushButton;
            BitmapImage iconCR = LoadImage("HMVTools.Resources.constraints_32.png");
            if (iconCR != null) btnCR.LargeImage = iconCR;





            // ── Annotation Tools Panel ──
            RibbonPanel panelAnnot = app.CreateRibbonPanel("HMV Tools", "Annotation Tools");


            PushButtonData btnAutoDim = new PushButtonData(
                "AutoDimension",
                "Auto\nDimension",
                path,
                "HMVTools.AutoDimensionCommand");
            btnAutoDim.ToolTip = "Automatically dimension intersecting walls along a selected line";
            btnAutoDim.LongDescription =
                "Creates a dimension segment along a selected Detail or Model Line, "
                + "automatically detecting and dimensioning intersecting walls.\n\n"
                + "Features:\n"
                + "• Interactive tolerance setting to skip close lines\n"
                + "• Auto-offset for text to prevent collisions on small dimensions\n"
                + "• Works with straight drawn lines acting as a dimension path";

            PushButton btnAD = panelAnnot.AddItem(btnAutoDim) as PushButton;

            // Note: Remember to add a 32x32 icon named 'autodim_32.png' to your project Resources!
            BitmapImage iconAD = LoadImage("HMVTools.Resources.autodim_32.png");
            if (iconAD != null) btnAD.LargeImage = iconAD;

            PushButtonData btnTopoLines = new PushButtonData(
                "TopoToLines",
                "Topo to\nLines",
                path,
                "HMVTools.TopographyToLinesCommand");
            btnTopoLines.ToolTip = "Convert topography to detail lines";
            btnTopoLines.LongDescription =
                "Extracts contour lines from a topography contained in a Revit link "
                + "and converts them into native detail lines in the active view.\n\n"
                + "Features:\n"
                + "• Select a Revit link containing topography\n"
                + "• Automatic intersection calculation with the view plane\n"
                + "• Detail line creation in the active view\n"
                + "• Report of lines created and skipped";
            PushButton btnTopo = panelAnnot.AddItem(btnTopoLines) as PushButton;
            BitmapImage iconTopo = LoadImage("HMVTools.Resources.topo_32.png");
            if (iconTopo != null) btnTopo.LargeImage = iconTopo;

            PushButtonData btnPipeAnnot = new PushButtonData(
                "PipeandFramingAnnotations",
                "Pipe, Framing\nAnnotations",
                path,
                "HMVTools.PlaceAnnotationsAlongPipeCmd");
            btnPipeAnnot.ToolTip = "Place annotations along paths";
            btnPipeAnnot.LongDescription =
                "Places instances of a generic annotation family along the path "
                + "of Pipes, Flex Pipes, and Structural Framing elements.\n\n"
                + "Features:\n"
                + "• Selection of Pipes, Flex Pipes, and Structural Framing\n"
                + "• Configurable spacing between annotations\n"
                + "• Automatic orientation based on path tangent\n"
                + "• Works with straight segments, arcs, and splines";
            PushButton btnPA = panelAnnot.AddItem(btnPipeAnnot) as PushButton;
            BitmapImage iconPipeAnnot = LoadImage("HMVTools.Resources.pipeannotation_32.png");
            if (iconPipeAnnot != null) btnPA.LargeImage = iconPipeAnnot;

            PushButtonData btnGridLevel = new PushButtonData(
                "GridLevelExtent",
                "Grid/Level\nExtent",
                path,
                "HMVTools.GridLevelExtentCommand");
            btnGridLevel.ToolTip = "Switch grids and levels between 2D and 3D extent";
            btnGridLevel.LongDescription =
                "Converts grids and levels between 2D (ViewSpecific) and 3D (Model) "
                + "datum extent across multiple views.\n\n"
                + "Features:\n"
                + "• Toggle between 2D (independent per view) and 3D (global)\n"
                + "• Choose to process Grids, Levels, or both\n"
                + "• Select individual views or apply to all views at once\n"
                + "• Summary report of all changes made";
            PushButton btnGL = panelAnnot.AddItem(btnGridLevel) as PushButton;
            BitmapImage iconGL = LoadImage("HMVTools.Resources.gridlevel_32.png");
            if (iconGL != null) btnGL.LargeImage = iconGL;

            PushButtonData btnSpotElev = new PushButtonData(
             "SpotElevOnFloor",
             "Spot Elev\non Floor",
              path,
            "HMVTools.SpotElevationCommand");
            btnSpotElev.ToolTip = "Place spot elevations on linked floor at foundation centers";
            btnSpotElev.LongDescription =
                "Places a Spot Elevation annotation at each selected foundation's center XY "
                + "position, reading the Z from the top face of a floor in a linked model.\n\n"
                + "Workflow:\n"
                + "1. Choose foundation source (active model or a link) and floor link.\n"
                + "2. Pick foundations in the canvas.\n"
                + "3. The tool ray-casts to the linked floor's top face and places "
                + "spot elevations with leaders.";
            PushButton btnSE = panelAnnot.AddItem(btnSpotElev) as PushButton;
            BitmapImage iconSE = LoadImage("HMVTools.Resources.spotelev_32.png");
            if (iconSE != null) btnSE.LargeImage = iconSE;

            PushButtonData btnAnnotTag = new PushButtonData(
            "GenericAnnotationTag",
            "Generic Annot\nTag",
            path,
            "HMVTools.GenericAnnotationTagCommand");
            btnAnnotTag.ToolTip = "Tag linked elements with a TextNote + leader";
            btnAnnotTag.LongDescription =
                "Reads a parameter value from elements in a Revit link "
                + "and places a Generic Annotation at each element's location "
                + "with the value written into the annotation's text parameter.\n\n"
                + "Workflow:\n"
                + "1. Pick elements in a linked model.\n"
                + "2. Choose annotation family, offset, and source parameter.\n"
                + "3. Annotations are placed with the parameter value.";
            PushButton btnAT = panelAnnot.AddItem(btnAnnotTag) as PushButton;
            BitmapImage iconAT = LoadImage("HMVTools.Resources.annottag_32.png");
            if (iconAT != null) btnAT.LargeImage = iconAT;

            PushButtonData btnAlignSpot = new PushButtonData(
                "AlignSpotElevations",
                "Align Spot\nElevations",
                path,
                "HMVTools.AlignSpotElevationsCommand");
            btnAlignSpot.ToolTip = "Align spot elevation text and leaders to a common line";
            btnAlignSpot.LongDescription =
                "Aligns the text position and leader shoulder of multiple spot "
                + "elevations to a single X or Y coordinate.\n\n"
                + "Features:\n"
                + "• Align to X-Axis (vertical stack) or Y-Axis (horizontal stack)\n"
                + "• Pick a reference line on screen or type a coordinate value\n"
                + "• Automatically adjusts leader shoulders to prevent broken angles\n"
                + "• Works with pre-selected spots or interactive pick";
            PushButton btnAS = panelAnnot.AddItem(btnAlignSpot) as PushButton;
            BitmapImage iconAS = LoadImage("HMVTools.Resources.alignspot_32.png");
            if (iconAS != null) btnAS.LargeImage = iconAS;

            // ── Audit Panel ──
            RibbonPanel panelAudit = app.CreateRibbonPanel("HMV Tools", "Audit");

            PushButtonData btnTextAudit = new PushButtonData(
                "TextAudit",
                "Text\nAudit",
                path,
                "HMVTools.TextAuditCommand");
            btnTextAudit.ToolTip = "Standardize all text styles and annotation tag families";
            btnTextAudit.LongDescription =
                "Standardize Texts:\n\n"
                + "Step 1 – Standardize Properties: normalizes Font (Arial), Width (1.0), "
                + "Bold (No), Background (Opaque) and Size (snap to 1.5/2/2.5/3/3.5 mm) "
                + "across all TextNoteTypes.\n\n"
                + "Step 2 – Merge Types: groups TextNoteTypes by size, renames the "
                + "keeper to HMV_General_Xmm Arial and reassigns all instances.\n\n"
                + "Step 3 – Purge: removes empty types left with no instances.\n\n"
                + "Step 4 – Tag Families: opens each annotation family, standardizes "
                + "the internal text properties and reloads the family into the project.\n\n"
                + "Displays a complete report of all changes when finished.";
            PushButton btnTA = panelAudit.AddItem(btnTextAudit) as PushButton;
            BitmapImage iconTA = LoadImage("HMVTools.Resources.textaudit_32.png");
            if (iconTA != null) btnTA.LargeImage = iconTA;

            // Dimension Audit button
            PushButtonData btnDimAudit = new PushButtonData(
                "DimAudit",
                "Dim\nAudit",
                path,
                "HMVTools.DimensionAuditCommand");
            btnDimAudit.ToolTip = "Standardize all linear dimension types to HMV naming";
            btnDimAudit.LongDescription =
                "Standardize Dimensions:\n\n"
                + "Step A – Standardize Properties: normalizes Font (Arial), Width (1.0), "
                + "Bold (No), Background (Opaque) and Size (snap to 1.5/2/2.5/3/3.5 mm) "
                + "across all linear DimensionTypes.\n\n"
                + "Step B – Merge Types: groups by unit (m/mm), text size, and decimal "
                + "places (2 or 3). Renames keepers to:\n"
                + "  HMV_Acotado Lineal [m|mm]_[size]mm Arial CIV.[XX|XXX]\n"
                + "and reassigns all instances.\n\n"
                + "Step C – Purge: removes empty duplicate types.\n\n"
                + "Displays a complete report of all changes when finished.";
            PushButton btnDA = panelAudit.AddItem(btnDimAudit) as PushButton;
            BitmapImage iconDA = LoadImage("HMVTools.Resources.dimaudit_32.png");
            if (iconDA != null) btnDA.LargeImage = iconDA;

            // View Audit button
            PushButtonData btnViewAudit = new PushButtonData(
                "ViewAudit", "View\nAudit", path,
                "HMVTools.ViewAuditCommand");
            btnViewAudit.ToolTip = "Audit and rename views placed on sheets";
            btnViewAudit.LongDescription =
                "Displays an interactive table of all views placed on sheets:\n\n"
                + "• View type, original name, editable new name, sheet(s)\n"
                + "• Bulk prefix addition and text cutting on selected views\n"
                + "• Duplicate name detection scoped to same view type (red highlight)\n"
                + "• Conflicting views are skipped when applying changes";
            PushButton btnVA = panelAudit.AddItem(btnViewAudit) as PushButton;
            BitmapImage iconVA = LoadImage("HMVTools.Resources.viewaudit_32.png");
            if (iconVA != null) btnVA.LargeImage = iconVA;

            // Sheet Audit button
            PushButtonData btnSheetAudit = new PushButtonData(
                "SheetAudit", "Sheet\nAudit", path,
                "HMVTools.SheetAuditCommand");
            btnSheetAudit.ToolTip = "Audit and rename sheet numbers and names";
            btnSheetAudit.LongDescription =
                "Displays an interactive table of all sheets in the project:\n\n"
                + "• Editable sheet number and sheet name columns\n"
                + "• Bulk prefix addition and text cutting for both fields\n"
                + "• Duplicate number detection (red highlight)\n"
                + "• Sheets with conflicting numbers are skipped when applying";
            PushButton btnSA = panelAudit.AddItem(btnSheetAudit) as PushButton;
            BitmapImage iconSA = LoadImage("HMVTools.Resources.sheetaudit_32.png");
            if (iconSA != null) btnSA.LargeImage = iconSA;

            // ── Family Audit button ──────────────────────────
            PushButtonData btnFamilyAudit = new PushButtonData(
                "FamilyAudit",
                "Family\nAudit",
                path,
                "HMVTools.AuditFamiliesCommand");
            btnFamilyAudit.ToolTip = "Audit family versions against ADC source files";
            btnFamilyAudit.LongDescription =
                "Scans all loaded editable families and compares their VersionGuid "
                + "against the source .rfa file in the Autodesk Desktop Connector "
                + "(ADC/ACCDocs) workspace.\n\n"
                + "Features:\n"
                + "• Automatic ADC file hydration for ProjFS \"Online Only\" files\n"
                + "• Two-tier path resolution: Extensible Storage → directory search\n"
                + "• GUID-based comparison (fast, no full document open)\n"
                + "• SHA-256 hash for mismatched families (bit-for-bit evidence)\n"
                + "• Colour-coded report with clipboard export\n"
                + "• Auto-stamps families with traceability metadata";
            PushButton btnFA = panelAudit.AddItem(btnFamilyAudit) as PushButton;
            btnFA.Enabled = false;
            BitmapImage iconFA = LoadImage("HMVTools.Resources.familyaudit_32.png");
            if (iconFA != null) btnFA.LargeImage = iconFA;

            // ── Linked GIS Files Audit button ──────────────────────────
            PushButtonData btnLinkAudit = new PushButtonData(
                "LinkAudit",
                "Link GIS\nAudit",
                path,
                "HMVTools.LinkedFilesAuditCommand");
            btnLinkAudit.ToolTip = "Audit GIS and Coordinate data for all linked Revit files";
            btnLinkAudit.LongDescription =
                  "Extracts coordinate and site information from linked models:\n\n"
                    + "• Link Name & Status\n"
                    + "• GIS Coordinate System Code\n"
                    + "• Latitude & Longitude\n"
                    + "• Site Name defined in the link\n"
                    + "• Angle to True North\n\n"
                    + "The data can be viewed in a table and exported to CSV.";
            PushButton btnLA = panelAudit.AddItem(btnLinkAudit) as PushButton;
            BitmapImage iconLA = LoadImage("HMVTools.Resources.linkaudit_32.png");
            if (iconLA != null) btnLA.LargeImage = iconLA;


            // ── Transfer Tools Panel ──
            RibbonPanel panelTrans = app.CreateRibbonPanel("HMV Tools", "Transfer Tools");


            PushButtonData btnMigrate = new PushButtonData(
                "MigratePlans",
                "Migrate\nElements",
                path,
                "HMVTools.MigrateElementsCommand");
            btnMigrate.ToolTip = "Migrate elements and views to another open project";
            btnMigrate.LongDescription =
                "Transfers selected model elements and their associated views "
                + "(Floor Plans, Sections, Drafting Views, Legends) from the "
                + "current project to another open project file.\n\n"
                + "Features:\n"
                + "• Shared Coordinates alignment (exact world position)\n"
                + "• View reconstruction with matching crop region & view range\n"
                + "• 2-step annotation migration (Revit 2023 compatible)\n"
                + "• Duplicate type conflict resolution (uses destination types)\n"
                + "• All-or-nothing TransactionGroup for data integrity\n\n"
                + "Workflow:\n"
                + "1. Open both source and target documents in Revit.\n"
                + "2. Select elements in the source document.\n"
                + "3. Run this command, pick the target doc, and select views.\n"
                + "4. Review the migration report.";
            PushButton btnMig = panelTrans.AddItem(btnMigrate) as PushButton;
            BitmapImage iconMig = LoadImage("HMVTools.Resources.migrate_32.png");
            if (iconMig != null) btnMig.LargeImage = iconMig;



            PushButtonData btnTransferUnitsData = new PushButtonData(
            "TransferUnits",
            "Transfer\nUnits",
            path,
            "HMVTools.TransferUnitsCommand");

            btnTransferUnitsData.ToolTip = "Batch transfer Project Units from a master file to multiple target files.";

            btnTransferUnitsData.LongDescription =
                "Automates the transfer of unit settings across multiple Revit models without opening them in the UI.\n\n"
                + "Features:\n"
                + "• Headless batch processing for maximum performance\n"
                + "• Transfers global settings (Decimal symbol, Digit grouping)\n"
                + "• Deep-copies FormatOptions using the modern ForgeTypeId system\n"
                + "• Safely handles workshared central models by detaching from central";

            // Replace 'panelMain' with the actual RibbonPanel variable you want to use
            PushButton btnTransferUnits = panelTrans.AddItem(btnTransferUnitsData) as PushButton;

            // Make sure to add a 32x32 icon to your Resources and set its Build Action to 'Embedded Resource'
            BitmapImage iconTransferUnits = LoadImage("HMVTools.Resources.transfer_units_32.png");
            if (iconTransferUnits != null) btnTransferUnits.LargeImage = iconTransferUnits;


            // ── Cloud Tools Panel ──
            RibbonPanel panelCloud = app.CreateRibbonPanel("HMV Tools", "Cloud");


            // ── Massive link files into revit ──────────────────────────
            PushButtonData btnLinkCloudFiles = new PushButtonData(
                "LinkFiles",
                "Link Cloud\nFiles",
                path,
                "HMVTools.BatchCloudLinkCommand");
            btnLinkAudit.ToolTip = "Massive Link Files from cloud services (ACC)";
            btnLinkAudit.LongDescription =
                  "Get your project ID from your model:\n\n"
                    + "• Show possible link files inside the project\n"
                    + "• Allow to set Shared or PBP type of link\n"
                    + "Linked all from cloud services";
            PushButton btnLCF = panelCloud.AddItem(btnLinkCloudFiles) as PushButton;
            BitmapImage iconLCF = LoadImage("HMVTools.Resources.linkcloud_32.png");
            if (iconLCF != null) btnLCF.LargeImage = iconLCF;


            // ── Electrical Tools Panel ──
            RibbonPanel panelElect = app.CreateRibbonPanel("HMV Tools", "Electrical");

            PushButtonData btnElectConect = new PushButtonData(
                "ElectricConect",
                "Electrical\n Connect",
                path,
                "HMVTools.ElectricalConnectionCommand");
            btnElectConect.ToolTip = "Generate the electrical connection from a path .dxf";
            btnElectConect.LongDescription =
                  "Extracts path and coordinate information from a .dxf file:\n\n"
                    + "• Select the .dxf file\n"
                    + "• Pick connector A and connector B\n"
                    + "• Select the flex pipe Type.\n"
                    + "• Put the .dxf name inside a parameter\n"
                    + "After this we have the option to refresh the path.";
            PushButton btnElc = panelElect.AddItem(btnElectConect) as PushButton;
            BitmapImage iconElc = LoadImage("HMVTools.Resources.electricalconect_32.png");
            if (iconElc != null) btnElc.LargeImage = iconElc;

            // ══════════════════════════════════════════════════════
            //  PROACTIVE TRACKING — Event subscriptions
            // ══════════════════════════════════════════════════════
            //
            // These events fire on the Revit main thread. We capture
            // metadata here but defer all DB modifications to an
            // ExternalEvent handler (safe Revit API context).
            //
            _stampHandler = new StampFamilyHandler();
            _stampEvent = ExternalEvent.Create(_stampHandler);

            app.ControlledApplication.FamilyLoadingIntoDocument +=
                OnFamilyLoading;
            app.ControlledApplication.FamilyLoadedIntoDocument +=
                OnFamilyLoaded;

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            // ── Clean up event subscriptions ──────────────────────
            try
            {
                app.ControlledApplication.FamilyLoadingIntoDocument -=
                    OnFamilyLoading;
                app.ControlledApplication.FamilyLoadedIntoDocument -=
                    OnFamilyLoaded;
            }
            catch { /* Revit is shutting down — swallow gracefully */ }

            return Result.Succeeded;
        }

        // ══════════════════════════════════════════════════════════
        //  EVENT HANDLERS
        // ══════════════════════════════════════════════════════════

        /// <summary>
        /// Fires BEFORE the family is loaded. Captures the source
        /// file path so we can read it in the Loaded handler
        /// (Revit 2023's FamilyLoadedIntoDocumentEventArgs does NOT
        /// expose the path).
        /// </summary>
        private void OnFamilyLoading(object sender,
            FamilyLoadingIntoDocumentEventArgs e)
        {
            if (!ProactiveTrackingEnabled) return;

            try
            {
                string familyPath = e.FamilyPath;
                string familyName = e.FamilyName;

                if (string.IsNullOrEmpty(familyPath)
                    || string.IsNullOrEmpty(familyName))
                    return;

                // Only track families coming from the ADC workspace
                if (!FamilyTraceabilityManager.IsAdcPath(familyPath))
                    return;

                // Recursion guard: if we're currently reloading this
                // family because of parameter injection, skip.
                if (_stampingInProgress.Contains(familyName))
                    return;

                _pendingPaths[familyName] = familyPath;
            }
            catch { /* Never crash Revit from an event handler */ }
        }

        /// <summary>
        /// Fires AFTER the family is loaded. Computes the SHA-256 hash
        /// (blocking but typically fast for .rfa files) and queues
        /// the Revit DB stamping work to the ExternalEvent handler.
        /// </summary>
        private void OnFamilyLoaded(object sender,
            FamilyLoadedIntoDocumentEventArgs e)
        {
            if (!ProactiveTrackingEnabled) return;

            try
            {
                // Only process successful loads (Revit 2023 does not
                // expose FamilyLoadedStatus; a valid NewFamilyId means success)
                if (e.NewFamilyId == null
                    || e.NewFamilyId == ElementId.InvalidElementId)
                    return;

                string familyName = e.FamilyName;

                // Recursion guard
                if (_stampingInProgress.Contains(familyName))
                    return;

                // Retrieve the path captured in the Loading handler
                if (!_pendingPaths.TryRemove(familyName, out string adcPath))
                    return;

                // Hydrate the file if needed (should already be hydrated
                // since Revit just loaded it, but be safe)
                if (!FamilyTraceabilityManager.EnsureHydrated(adcPath))
                    return;

                // Compute SHA-256 while we're still in the event.
                // For typical .rfa files (< 20 MB) this takes < 500 ms.
                string hash = FamilyTraceabilityManager.ComputeSha256(adcPath);

                // Queue the Revit DB modifications
                _stampHandler.Enqueue(new StampRequest
                {
                    FamilyName = familyName,
                    FamilyId = e.NewFamilyId,
                    AdcPath = adcPath,
                    Sha256Hash = hash ?? "",
                    Timestamp = DateTime.UtcNow.ToString("o"),
                    UserMachineId = FamilyTraceabilityManager.GetUserMachineId(),
                    InjectParam = InjectParameterOnLoad
                });
                _stampEvent.Raise();
            }
            catch { /* Never crash Revit from an event handler */ }
        }

        // ══════════════════════════════════════════════════════════
        //  IMAGE LOADER
        // ══════════════════════════════════════════════════════════

        private BitmapImage LoadImage(string resourceName)
        {
            try
            {
                Stream stream = Assembly.GetExecutingAssembly()
                    .GetManifestResourceStream(resourceName);
                if (stream == null) return null;
                BitmapImage image = new BitmapImage();
                image.BeginInit();
                image.StreamSource = stream;
                image.EndInit();
                return image;
            }
            catch { return null; }
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  Data class for queued stamp operations
    // ═══════════════════════════════════════════════════════════════
    public class StampRequest
    {
        public string FamilyName { get; set; }
        public ElementId FamilyId { get; set; }
        public string AdcPath { get; set; }
        public string Sha256Hash { get; set; }
        public string Timestamp { get; set; }
        public string UserMachineId { get; set; }
        public bool InjectParam { get; set; }
    }

    // ═══════════════════════════════════════════════════════════════
    //  ExternalEvent handler — all Revit DB writes happen here
    // ═══════════════════════════════════════════════════════════════

    /// <summary>
    /// Processes queued <see cref="StampRequest"/>s in a valid Revit API
    /// context. Writes Extensible Storage and optionally injects the
    /// visible ADC_Library_Path parameter into each tracked family.
    /// </summary>
    public class StampFamilyHandler : IExternalEventHandler
    {
        private readonly Queue<StampRequest> _queue = new Queue<StampRequest>();
        private readonly object _lock = new object();

        public void Enqueue(StampRequest req)
        {
            lock (_lock) { _queue.Enqueue(req); }
        }

        public void Execute(UIApplication uiApp)
        {
            // Drain the queue
            List<StampRequest> batch;
            lock (_lock)
            {
                batch = new List<StampRequest>(_queue);
                _queue.Clear();
            }
            if (batch.Count == 0) return;

            Document doc = uiApp.ActiveUIDocument?.Document;
            if (doc == null) return;

            // ── 1. Write Extensible Storage (lightweight) ─────────
            using (Transaction tx = new Transaction(doc,
                "HMV - Stamp Family Traceability"))
            {
                tx.Start();
                try
                {
                    foreach (StampRequest req in batch)
                    {
                        Element el = doc.GetElement(req.FamilyId);
                        Family fam = el as Family;
                        if (fam == null) continue;

                        FamilyTraceabilityManager.WriteTraceData(fam,
                            new TraceData
                            {
                                AdcPath = req.AdcPath,
                                Sha256Hash = req.Sha256Hash,
                                LoadTimestamp = req.Timestamp,
                                UserMachineId = req.UserMachineId
                            });
                    }
                    tx.Commit();
                }
                catch
                {
                    if (tx.HasStarted()) tx.RollBack();
                }
            }

            // ── 2. Parameter injection (heavyweight, optional) ────
            foreach (StampRequest req in batch)
            {
                if (!req.InjectParam) continue;

                Element el = doc.GetElement(req.FamilyId);
                Family fam = el as Family;
                if (fam == null || fam.IsInPlace) continue;

                // Recursion guard: mark this family so the Loading
                // event ignores the reload triggered by EditFamily.
                App.StampingInProgress.Add(req.FamilyName);
                try
                {
                    FamilyTraceabilityManager.InjectVisibleParameter(
                        doc, fam, req.AdcPath);
                }
                catch { /* Log but never crash */ }
                finally
                {
                    App.StampingInProgress.Remove(req.FamilyName);
                }
            }
        }

        public string GetName() => "HMV Family Traceability Stamp Handler";
    }

}