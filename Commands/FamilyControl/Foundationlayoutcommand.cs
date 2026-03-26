using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using Color = System.Windows.Media.Color;
using Button = System.Windows.Controls.Button;
using ComboBox = System.Windows.Controls.ComboBox;
using Grid = System.Windows.Controls.Grid;

namespace HMVTools
{
    // ═══════════════════════════════════════════════════════════════
    //  Plain data class — no Revit references, safe for WPF window
    // ═══════════════════════════════════════════════════════════════

    public class FoundationItemData
    {
        public int ElementIdInt;
        public string Name;

        // Plan extents in meters (for canvas drawing)
        public double PlanMinXm, PlanMinYm, PlanMaxXm, PlanMaxYm;

        // NTCE = BBox top elevation
        public double NtceSurveyM;
        public double NtceProjectM;

        // NAP = max floor Z from raycasts
        public bool HasNap;
        public double NapSurveyM;
        public double NapProjectM;

        // For Apply delta computation (internal feet)
        public double OrigTopZFeet;

        // Editing
        public bool CanEdit;         // = HasNap
        public double NewNtceSurveyM; // user-edited value
    }

    // ═══════════════════════════════════════════════════════════════
    //  Selection filter for structural foundations
    // ═══════════════════════════════════════════════════════════════

    public class FoundationSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem?.Category?.Id.IntegerValue
                == (int)BuiltInCategory.OST_StructuralFoundation;
        }
        public bool AllowReference(Reference reference, XYZ position) => false;
    }

    // ═══════════════════════════════════════════════════════════════
    //  Pre-dialog: Link + 3D View selector
    // ═══════════════════════════════════════════════════════════════

    public class FoundationLinkSelectWindow : Window
    {
        public int SelectedLinkIndex { get; private set; } = -1;
        public int SelectedView3DIndex { get; private set; } = -1;

        static readonly Color COL_BG = Color.FromRgb(245, 245, 248);
        static readonly Color COL_ACCENT = Color.FromRgb(0, 120, 212);
        static readonly Color COL_TEXT = Color.FromRgb(30, 30, 30);
        static readonly Color COL_SUB = Color.FromRgb(120, 120, 130);
        static readonly Color COL_BTN = Color.FromRgb(240, 240, 243);
        static readonly Color COL_BORDER = Color.FromRgb(200, 200, 210);

        public FoundationLinkSelectWindow(
            List<string> linkNames,
            List<string> viewNames,
            int foundationCount)
        {
            Title = "HMV Tools — Foundation Control";
            Width = 480; Height = 320;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(COL_BG);

            StackPanel root = new StackPanel { Margin = new Thickness(24) };
            Content = root;

            root.Children.Add(new TextBlock
            {
                Text = "Foundation Control",
                FontSize = 20,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(COL_TEXT),
                Margin = new Thickness(0, 0, 0, 4)
            });
            root.Children.Add(new TextBlock
            {
                Text = $"{foundationCount} foundation(s) selected. Choose the floor link and a 3D view for raycasting.",
                FontSize = 11,
                Foreground = new SolidColorBrush(COL_SUB),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 18)
            });

            // Floor link combo
            root.Children.Add(new TextBlock
            {
                Text = "Floor Link (NAP source):",
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(COL_ACCENT),
                Margin = new Thickness(0, 0, 0, 4)
            });
            ComboBox cmbLink = new ComboBox
            {
                Height = 32,
                FontSize = 12,
                VerticalContentAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 14)
            };
            int autoIdx = -1;
            for (int i = 0; i < linkNames.Count; i++)
            {
                cmbLink.Items.Add(linkNames[i]);
                if (autoIdx < 0 && linkNames[i].Contains("303"))
                    autoIdx = i;
            }
            cmbLink.SelectedIndex = autoIdx >= 0 ? autoIdx : 0;
            root.Children.Add(cmbLink);

            // 3D view combo
            root.Children.Add(new TextBlock
            {
                Text = "3D View (for ReferenceIntersector):",
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(COL_ACCENT),
                Margin = new Thickness(0, 0, 0, 4)
            });
            ComboBox cmbView = new ComboBox
            {
                Height = 32,
                FontSize = 12,
                VerticalContentAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 20)
            };
            foreach (string v in viewNames) cmbView.Items.Add(v);
            cmbView.SelectedIndex = 0;
            root.Children.Add(cmbView);

            // Buttons
            StackPanel btnRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            Button btnCancel = MkBtn("Cancel", COL_BTN, COL_TEXT, 100);
            btnCancel.Click += (s, e) => { DialogResult = false; Close(); };
            btnRow.Children.Add(btnCancel);

            Button btnOk = MkBtn("Continue →", COL_ACCENT, Colors.White, 130);
            btnOk.FontWeight = FontWeights.SemiBold;
            btnOk.Click += (s, e) =>
            {
                SelectedLinkIndex = cmbLink.SelectedIndex;
                SelectedView3DIndex = cmbView.SelectedIndex;
                DialogResult = true; Close();
            };
            btnRow.Children.Add(btnOk);
            root.Children.Add(btnRow);
        }

        Button MkBtn(string t, Color bg, Color fg, double w)
        {
            Button b = new Button
            {
                Content = t,
                Width = w,
                Height = 34,
                FontSize = 12,
                Margin = new Thickness(0, 0, 8, 0),
                Cursor = Cursors.Hand,
                Foreground = new SolidColorBrush(fg),
                Background = new SolidColorBrush(bg),
                BorderThickness = new Thickness(0)
            };
            var tp = new ControlTemplate(typeof(Button));
            var bd = new FrameworkElementFactory(typeof(Border));
            bd.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
            bd.SetValue(Border.BackgroundProperty, new SolidColorBrush(bg));
            bd.SetValue(Border.PaddingProperty, new Thickness(12, 4, 12, 4));
            var cp = new FrameworkElementFactory(typeof(ContentPresenter));
            cp.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            cp.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            bd.AppendChild(cp); tp.VisualTree = bd; b.Template = tp;
            return b;
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  Main Command
    // ═══════════════════════════════════════════════════════════════

    [Transaction(TransactionMode.Manual)]
    public class FoundationLayoutCommand : IExternalCommand
    {
        private const double FeetToM = 0.3048;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // ── 1. Collect foundations from active selection ──────
            ICollection<ElementId> selIds = uidoc.Selection.GetElementIds();
            List<FamilyInstance> foundations = new List<FamilyInstance>();
            foreach (ElementId id in selIds)
            {
                Element el = doc.GetElement(id);
                if (el == null) continue;
                if (el.Category?.Id.IntegerValue
                    != (int)BuiltInCategory.OST_StructuralFoundation) continue;
                if (el is FamilyInstance fi) foundations.Add(fi);
            }

            if (foundations.Count == 0)
            {
                // Prompt user to select
                try
                {
                    IList<Reference> picked = uidoc.Selection.PickObjects(
                        ObjectType.Element,
                        new FoundationSelectionFilter(),
                        "Select structural foundations. Finish when done.");

                    foreach (Reference r in picked)
                    {
                        Element el = doc.GetElement(r);
                        if (el is FamilyInstance fi) foundations.Add(fi);
                    }
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                { return Result.Cancelled; }
            }

            if (foundations.Count == 0)
            {
                TaskDialog.Show("HMV Tools",
                    "No structural foundations found in selection.");
                return Result.Cancelled;
            }

            // ── 2. Collect Revit links ───────────────────────────
            var linkInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(li => li.GetLinkDocument() != null)
                .ToList();

            if (linkInstances.Count == 0)
            {
                TaskDialog.Show("HMV Tools",
                    "No loaded Revit links found. A link with floors is required for NAP raycasting.");
                return Result.Cancelled;
            }

            var linkNames = linkInstances.Select(li => li.Name).ToList();

            // ── 3. Collect valid 3D views ────────────────────────
            var views3D = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .Where(v => !v.IsTemplate && !v.IsSectionBoxActive)
                .OrderBy(v => v.Name)
                .ToList();

            if (views3D.Count == 0)
            {
                TaskDialog.Show("HMV Tools",
                    "No valid 3D views found (need at least one without an active section box).");
                return Result.Cancelled;
            }

            var viewNames = views3D.Select(v => v.Name).ToList();

            // ── 4. Show pre-dialog ───────────────────────────────
            var preWin = new FoundationLinkSelectWindow(
                linkNames, viewNames, foundations.Count);
            if (preWin.ShowDialog() != true) return Result.Cancelled;

            RevitLinkInstance floorLink = linkInstances[preWin.SelectedLinkIndex];
            View3D view3d = views3D[preWin.SelectedView3DIndex];

            // ── 5. Survey offset ─────────────────────────────────
            ProjectPosition pp = doc.ActiveProjectLocation
                .GetProjectPosition(XYZ.Zero);
            double surveyOffsetFeet = pp.Elevation;

            // ── 6. Build ReferenceIntersectors ───────────────────

            // Floor intersector (for NAP — linked floors)
            var floorRI = new ReferenceIntersector(
                new ElementCategoryFilter(BuiltInCategory.OST_Floors),
                FindReferenceTarget.Face, view3d);
            floorRI.FindReferencesInRevitLinks = true;

            // Foundation intersector (for NTCE — host element top face)
            // BoundingBox.Max.Z is unreliable: it includes anchor bolts,
            // rebar hooks, sub-components above the actual concrete top.
            // Raycast finds the true top face, matching spot elevation readings.
            var fndCats = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_StructuralFoundation,
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_GenericModel
            };
            var fndRI = new ReferenceIntersector(
                new ElementMulticategoryFilter(fndCats),
                FindReferenceTarget.Face, view3d);
            fndRI.FindReferencesInRevitLinks = false; // host model only

            ElementId floorLinkId = floorLink.Id;

            // ── 7. Process each foundation ───────────────────────
            var items = new List<FoundationItemData>();
            var debugInfo = new List<string>();
            int skipped = 0;

            foreach (FamilyInstance fi in foundations)
            {
                BoundingBoxXYZ bb = fi.get_BoundingBox(null);
                if (bb == null) { skipped++; continue; }

                LocationPoint lp = fi.Location as LocationPoint;
                if (lp == null) { skipped++; continue; }

                string mark = fi.get_Parameter(
                    BuiltInParameter.ALL_MODEL_MARK)?.AsString() ?? "";
                string name = string.IsNullOrEmpty(mark)
                    ? $"{fi.Name} (Id {fi.Id.IntegerValue})"
                    : $"{fi.Name} [{mark}] (Id {fi.Id.IntegerValue})";

                double topZFeet = bb.Max.Z; // fallback

                // NTCE: raycast to find the true top face of the
                // foundation (bb.Max.Z includes anchor bolts etc.)
                double? rayTopZ = FindTopFaceZ(
                    fndRI, bb.Min.X, bb.Max.X,
                    bb.Min.Y, bb.Max.Y, topZFeet,
                    fi.Id, debugInfo);
                if (rayTopZ.HasValue)
                    topZFeet = rayTopZ.Value;

                // NTCE in survey and project meters
                double ntceSurveyM = (topZFeet + surveyOffsetFeet) * FeetToM;
                double ntceProjectM = topZFeet * FeetToM;

                // Plan extents in meters
                double planMinXm = bb.Min.X * FeetToM;
                double planMinYm = bb.Min.Y * FeetToM;
                double planMaxXm = bb.Max.X * FeetToM;
                double planMaxYm = bb.Max.Y * FeetToM;

                // NAP: dense downward raycast
                double? maxFloorZ = FindMaxFloorZ(
                    floorRI, bb.Min.X, bb.Max.X,
                    bb.Min.Y, bb.Max.Y, topZFeet,
                    floorLinkId, debugInfo);

                bool hasNap = maxFloorZ.HasValue;
                double napSurveyM = hasNap
                    ? (maxFloorZ.Value + surveyOffsetFeet) * FeetToM : 0;
                double napProjectM = hasNap
                    ? maxFloorZ.Value * FeetToM : 0;

                items.Add(new FoundationItemData
                {
                    ElementIdInt = fi.Id.IntegerValue,
                    Name = name,
                    PlanMinXm = planMinXm,
                    PlanMinYm = planMinYm,
                    PlanMaxXm = planMaxXm,
                    PlanMaxYm = planMaxYm,
                    NtceSurveyM = ntceSurveyM,
                    NtceProjectM = ntceProjectM,
                    HasNap = hasNap,
                    NapSurveyM = napSurveyM,
                    NapProjectM = napProjectM,
                    OrigTopZFeet = topZFeet,
                    CanEdit = hasNap,
                    NewNtceSurveyM = ntceSurveyM
                });
            }

            if (items.Count == 0)
            {
                TaskDialog.Show("HMV Tools",
                    "No valid foundations to process (all skipped due to missing BoundingBox or LocationPoint).");
                return Result.Cancelled;
            }

            // ── 8. Show main window ──────────────────────────────
            string linkLabel = floorLink.Name;
            double surveyOffsetM = surveyOffsetFeet * FeetToM;

            var mainWin = new FoundationLayoutWindow(
                items, linkLabel, surveyOffsetM);

            if (mainWin.ShowDialog() != true) return Result.Cancelled;

            List<FoundationItemData> results = mainWin.Items;

            // ── 9. Apply changes via Transaction ─────────────────
            int moved = 0;
            using (Transaction tx = new Transaction(doc,
                "HMV — Foundation Elevation Adjust"))
            {
                tx.Start();
                foreach (FoundationItemData item in results)
                {
                    if (!item.CanEdit) continue;
                    double deltaSurveyM = item.NewNtceSurveyM - item.NtceSurveyM;
                    if (Math.Abs(deltaSurveyM) < 0.0001) continue;

                    double deltaFeet = deltaSurveyM / FeetToM;

                    ElementId eid = new ElementId(item.ElementIdInt);
                    Element el = doc.GetElement(eid);
                    if (el == null) continue;

                    LocationPoint loc = el.Location as LocationPoint;
                    if (loc == null) continue;

                    XYZ pt = loc.Point;
                    loc.Point = new XYZ(pt.X, pt.Y, pt.Z + deltaFeet);
                    moved++;
                }
                tx.Commit();
            }

            // ── 10. Report ───────────────────────────────────────
            string report =
                "FOUNDATION CONTROL SUMMARY\n"
                + "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
                + $"Foundations processed: {items.Count}\n"
                + $"Elevations adjusted: {moved}\n";
            if (skipped > 0)
                report += $"Skipped (no BBox/LocationPoint): {skipped}\n";
            report += $"\nSurvey offset: {surveyOffsetM:F3} m\n"
                + $"Floor link: {linkLabel}\n"
                + $"3D View: {view3d.Name}";

            TaskDialog.Show("HMV Tools — Foundation Control", report);
            return Result.Succeeded;
        }

        // ═════════════════════════════════════════════════════════
        //  Dense downward raycast — returns max floor Z (feet)
        // ═════════════════════════════════════════════════════════

        private double? FindMaxFloorZ(
            ReferenceIntersector floorRI,
            double xMin, double xMax,
            double yMin, double yMax,
            double zTop,
            ElementId linkId,
            List<string> debugInfo)
        {
            double dx = xMax - xMin;
            double dy = yMax - yMin;
            double mx = (xMin + xMax) / 2.0;
            double my = (yMin + yMax) / 2.0;

            var samples = new List<XYZ>();

            // 4 corners
            samples.Add(new XYZ(xMin, yMin, 0));
            samples.Add(new XYZ(xMax, yMin, 0));
            samples.Add(new XYZ(xMin, yMax, 0));
            samples.Add(new XYZ(xMax, yMax, 0));

            // 4 edge midpoints
            samples.Add(new XYZ(mx, yMin, 0));
            samples.Add(new XYZ(mx, yMax, 0));
            samples.Add(new XYZ(xMin, my, 0));
            samples.Add(new XYZ(xMax, my, 0));

            // Center
            samples.Add(new XYZ(mx, my, 0));

            // 5×5 interior grid (inset 10 % from edges)
            for (int ix = 0; ix < 5; ix++)
                for (int iy = 0; iy < 5; iy++)
                    samples.Add(new XYZ(
                        xMin + dx * (0.1 + 0.8 * ix / 4.0),
                        yMin + dy * (0.1 + 0.8 * iy / 4.0), 0));

            // Inset corner samples (feet offsets)
            double[] ins = { 0.05, 0.1, 0.2, 0.33, 0.5 };
            foreach (double i in ins)
            {
                samples.Add(new XYZ(xMin + i, yMin + i, 0));
                samples.Add(new XYZ(xMax - i, yMin + i, 0));
                samples.Add(new XYZ(xMin + i, yMax - i, 0));
                samples.Add(new XYZ(xMax - i, yMax - i, 0));
            }

            // Quarter points along edges
            for (int q = 1; q <= 3; q++)
            {
                double fx = xMin + dx * q / 4.0;
                double fy = yMin + dy * q / 4.0;
                samples.Add(new XYZ(fx, yMin, 0));
                samples.Add(new XYZ(fx, yMax, 0));
                samples.Add(new XYZ(xMin, fy, 0));
                samples.Add(new XYZ(xMax, fy, 0));
            }

            double rayOriginZ = zTop + 200;
            double maxZ = double.MinValue;
            bool found = false;
            int totalHits = 0, linkHits = 0;

            foreach (XYZ xy in samples)
            {
                XYZ origin = new XYZ(xy.X, xy.Y, rayOriginZ);
                var hits = floorRI.Find(origin, XYZ.BasisZ.Negate());
                if (hits == null) continue;

                foreach (var rwc in hits)
                {
                    totalHits++;
                    Reference r = rwc.GetReference();
                    if (r.ElementId != linkId) continue;
                    linkHits++;

                    double hitZ = r.GlobalPoint != null
                        ? r.GlobalPoint.Z
                        : (rayOriginZ - rwc.Proximity);

                    if (hitZ > maxZ) { maxZ = hitZ; found = true; }
                }
            }

            debugInfo.Add($"  rays={samples.Count}, total={totalHits}, link={linkHits}, found={found}");
            return found ? (double?)maxZ : null;
        }

        // ═════════════════════════════════════════════════════════
        //  Raycast downward to find the true top face of a
        //  host-model foundation. BoundingBox.Max.Z is unreliable
        //  because it includes anchor bolts, rebar, etc.
        //  Returns the highest Z (closest to ray origin) in feet.
        // ═════════════════════════════════════════════════════════

        private double? FindTopFaceZ(
            ReferenceIntersector fndRI,
            double xMin, double xMax,
            double yMin, double yMax,
            double bbTopZ,
            ElementId targetId,
            List<string> debugInfo)
        {
            double dx = xMax - xMin;
            double dy = yMax - yMin;
            double mx = (xMin + xMax) / 2.0;
            double my = (yMin + yMax) / 2.0;

            var samples = new List<XYZ>();

            // 7×7 grid (same density as SpotElevationCommand)
            for (int ix = 0; ix < 7; ix++)
                for (int iy = 0; iy < 7; iy++)
                    samples.Add(new XYZ(
                        xMin + dx * (0.05 + 0.9 * ix / 6.0),
                        yMin + dy * (0.05 + 0.9 * iy / 6.0), 0));

            // 4 corners + edge midpoints + center
            samples.Add(new XYZ(xMin, yMin, 0));
            samples.Add(new XYZ(xMax, yMin, 0));
            samples.Add(new XYZ(xMin, yMax, 0));
            samples.Add(new XYZ(xMax, yMax, 0));
            samples.Add(new XYZ(mx, yMin, 0));
            samples.Add(new XYZ(mx, yMax, 0));
            samples.Add(new XYZ(xMin, my, 0));
            samples.Add(new XYZ(xMax, my, 0));
            samples.Add(new XYZ(mx, my, 0));

            // Inset corner probes
            double[] ins = { 0.05, 0.1, 0.2, 0.33, 0.5 };
            foreach (double i in ins)
            {
                samples.Add(new XYZ(xMin + i, yMin + i, 0));
                samples.Add(new XYZ(xMax - i, yMin + i, 0));
                samples.Add(new XYZ(xMin + i, yMax - i, 0));
                samples.Add(new XYZ(xMax - i, yMax - i, 0));
            }

            double rayOriginZ = bbTopZ + 200;
            int totalHits = 0, matchHits = 0;

            // Collect all Z values from matching hits
            var hitZValues = new List<double>();

            foreach (XYZ xy in samples)
            {
                var hits = fndRI.Find(
                    new XYZ(xy.X, xy.Y, rayOriginZ),
                    XYZ.BasisZ.Negate());
                if (hits == null) continue;

                // For each ray, take only the FIRST hit matching our element
                // (closest to ray origin = highest face at this XY)
                double bestProxThisRay = double.MaxValue;
                double bestZThisRay = 0;
                bool foundThisRay = false;

                foreach (var rwc in hits)
                {
                    totalHits++;
                    Reference r = rwc.GetReference();
                    if (r.ElementId != targetId) continue;
                    matchHits++;

                    if (rwc.Proximity < bestProxThisRay)
                    {
                        bestProxThisRay = rwc.Proximity;
                        bestZThisRay = r.GlobalPoint != null
                            ? r.GlobalPoint.Z
                            : (rayOriginZ - rwc.Proximity);
                        foundThisRay = true;
                    }
                }

                if (foundThisRay)
                    hitZValues.Add(bestZThisRay);
            }

            debugInfo.Add($"  NTCE-ray: rays={samples.Count}, total={totalHits}, match={matchHits}, zValues={hitZValues.Count}");

            if (hitZValues.Count == 0) return null;

            // Vote-by-frequency: group Z values by tolerance (5 mm ≈ 0.016 ft).
            // The concrete top face is large → hit by MANY rays.
            // Anchor bolt tops are tiny → hit by very FEW rays.
            // Pick the Z group with the most votes.
            double groupTol = 0.017; // ~5 mm in feet
            var groups = new List<KeyValuePair<double, int>>(); // avgZ, count

            foreach (double z in hitZValues)
            {
                bool merged = false;
                for (int gi = 0; gi < groups.Count; gi++)
                {
                    if (Math.Abs(z - groups[gi].Key) < groupTol)
                    {
                        // Running average + increment count
                        int newCount = groups[gi].Value + 1;
                        double newAvg = groups[gi].Key
                            + (z - groups[gi].Key) / newCount;
                        groups[gi] = new KeyValuePair<double, int>(
                            newAvg, newCount);
                        merged = true;
                        break;
                    }
                }
                if (!merged)
                    groups.Add(new KeyValuePair<double, int>(z, 1));
            }

            // Find group with maximum votes
            var winner = groups.OrderByDescending(g => g.Value).First();
            debugInfo.Add($"  NTCE-vote: groups={groups.Count}, winner Z={winner.Key * FeetToM:F4}m, votes={winner.Value}/{hitZValues.Count}");

            return winner.Key;
        }
    }
}