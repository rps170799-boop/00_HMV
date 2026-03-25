using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using Color = System.Windows.Media.Color;
using TextBox = System.Windows.Controls.TextBox;
using Button = System.Windows.Controls.Button;
using Grid = System.Windows.Controls.Grid;

namespace HMVTools
{
    // ═══════════════════════════════════════════════════════════════
    //  Plain data class for audit results — no Revit references,
    //  safe to pass to WPF window across boundary.
    // ═══════════════════════════════════════════════════════════════
    public class FamilyAuditItem
    {
        public string FamilyName { get; set; } = "";
        public string Category { get; set; } = "";
        public string AdcPath { get; set; } = "";
        public string Status { get; set; } = "";   // Match | Mismatch | No ADC Path | Error
        public string LoadedGuid { get; set; } = "";
        public string SourceGuid { get; set; } = "";
        public string Sha256Hash { get; set; } = "";
        public string ErrorDetail { get; set; } = "";
    }

    // ═══════════════════════════════════════════════════════════════
    //  AUDIT COMMAND
    // ═══════════════════════════════════════════════════════════════

    /// <summary>
    /// Scans loaded editable families in the active document and
    /// compares each one's VersionGuid against the source .rfa file
    /// in the ADC workspace. Shows a config window first, then a
    /// colour-coded report.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class AuditFamiliesCommand : IExternalCommand
    {
        /// <summary>Set per-run from the config window.</summary>
        private string _adcSearchRoot;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // ── 0. Gather context for the config window ───────────
                List<Family> allFamilies = new FilteredElementCollector(doc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .Where(f => f.IsEditable && !f.IsInPlace)
                    .ToList();

                if (allFamilies.Count == 0)
                {
                    TaskDialog.Show("HMV - Family Audit",
                        "No editable families found in the active document.");
                    return Result.Succeeded;
                }

                // Distinct category names for the "By Category" filter
                List<string> categories = allFamilies
                    .Select(f => f.FamilyCategory?.Name ?? "(none)")
                    .Distinct()
                    .ToList();

                // Count pre-selected family instances (for "Selected" scope)
                ICollection<ElementId> selIds = uidoc.Selection.GetElementIds();
                HashSet<ElementId> selectedFamilyIds = new HashSet<ElementId>();
                foreach (ElementId id in selIds)
                {
                    Element el = doc.GetElement(id);
                    if (el is FamilyInstance fi)
                        selectedFamilyIds.Add(fi.Symbol.Family.Id);
                    else if (el is Family fam)
                        selectedFamilyIds.Add(fam.Id);
                }

                // ── 1. Show config window ─────────────────────────────
                FamilyAuditConfigWindow configWin =
                    new FamilyAuditConfigWindow(categories, selectedFamilyIds.Count);

                if (configWin.ShowDialog() != true)
                    return Result.Cancelled;

                FamilyAuditConfig cfg = configWin.ResultConfig;
                _adcSearchRoot = cfg.AdcRootPath;

                // Apply proactive tracking toggle
                App.ProactiveTrackingEnabled = cfg.EnableProactiveTracking;

                // ── 2. Filter families by scope ───────────────────────
                List<Family> families;
                switch (cfg.Scope)
                {
                    case AuditScope.SelectedFamilies:
                        families = allFamilies
                            .Where(f => selectedFamilyIds.Contains(f.Id))
                            .ToList();
                        break;

                    case AuditScope.ByCategory:
                        families = allFamilies
                            .Where(f => (f.FamilyCategory?.Name ?? "(none)") == cfg.CategoryFilter)
                            .ToList();
                        break;

                    default:  // AllFamilies
                        families = allFamilies;
                        break;
                }

                if (families.Count == 0)
                {
                    TaskDialog.Show("HMV - Family Audit",
                        "No families match the selected scope.");
                    return Result.Succeeded;
                }

                // ── 3. Audit each family ──────────────────────────────
                List<FamilyAuditItem> results = new List<FamilyAuditItem>();

                foreach (Family fam in families)
                {
                    FamilyAuditItem item = new FamilyAuditItem
                    {
                        FamilyName = fam.Name,
                        Category = fam.FamilyCategory?.Name ?? "(none)"
                    };

                    try
                    {
                        // ── 3a. Resolve ADC path ──────────────────────
                        string adcPath = GetAdcPathForFamily(doc, fam);
                        item.AdcPath = adcPath ?? "";

                        if (string.IsNullOrEmpty(adcPath))
                        {
                            item.Status = "No ADC Path";
                            results.Add(item);
                            continue;
                        }

                        // ── 3b. Hydrate if ProjFS-offline ────────────
                        if (!FamilyTraceabilityManager.EnsureHydrated(adcPath))
                        {
                            item.Status = "Error";
                            item.ErrorDetail = "Hydration failed or file not found.";
                            results.Add(item);
                            continue;
                        }

                        // ── 3c. Extract source GUID via BasicFileInfo ─
                        //
                        // BasicFileInfo.Extract is FAST — it reads only
                        // the file header, never opens the full document.
                        //
                        // API casing note:
                        //   Source : BasicFileInfo → .GetDocumentVersion().VersionGUID
                        //   Loaded : Family.VersionGuid
                        // The case mismatch is intentional in the Revit API.
                        //
                        BasicFileInfo info = BasicFileInfo.Extract(adcPath);
                        Guid sourceGuid = info.GetDocumentVersion().VersionGUID;
                        Guid loadedGuid = fam.VersionGuid;

                        item.SourceGuid = sourceGuid.ToString("D");
                        item.LoadedGuid = loadedGuid.ToString("D");

                        item.Status = (sourceGuid == loadedGuid) ? "Match" : "Mismatch";

                        // ── 3d. Compute SHA-256 for mismatches (extra evidence) ─
                        if (item.Status == "Mismatch")
                        {
                            item.Sha256Hash =
                                FamilyTraceabilityManager.ComputeSha256(adcPath) ?? "(hash error)";
                        }
                    }
                    catch (Exception ex)
                    {
                        item.Status = "Error";
                        item.ErrorDetail = ex.Message;
                    }

                    results.Add(item);
                }

                // ── 4. Stamp families with ExtensibleStorage ──────────
                int stamped = 0;
                using (Transaction tx = new Transaction(doc, "HMV - Stamp Audit Results"))
                {
                    tx.Start();
                    try
                    {
                        foreach (Family fam in families)
                        {
                            FamilyAuditItem item = results.FirstOrDefault(
                                r => r.FamilyName == fam.Name);

                            if (item == null || string.IsNullOrEmpty(item.AdcPath))
                                continue;

                            TraceData existing =
                                FamilyTraceabilityManager.ReadTraceData(fam);

                            if (existing == null || existing.AdcPath != item.AdcPath)
                            {
                                FamilyTraceabilityManager.WriteTraceData(fam,
                                    new TraceData
                                    {
                                        AdcPath = item.AdcPath,
                                        Sha256Hash = item.Sha256Hash ?? "",
                                        LoadTimestamp = DateTime.UtcNow.ToString("o"),
                                        UserMachineId = FamilyTraceabilityManager.GetUserMachineId()
                                    });
                                stamped++;
                            }
                        }
                        tx.Commit();
                    }
                    catch
                    {
                        if (tx.HasStarted()) tx.RollBack();
                    }
                }

                // ── 5. Optional: inject visible parameter ─────────────
                int injected = 0;
                if (cfg.InjectParameter)
                {
                    foreach (Family fam in families)
                    {
                        FamilyAuditItem item = results.FirstOrDefault(
                            r => r.FamilyName == fam.Name);

                        if (item == null || string.IsNullOrEmpty(item.AdcPath))
                            continue;

                        try
                        {
                            if (FamilyTraceabilityManager.InjectVisibleParameter(
                                doc, fam, item.AdcPath))
                                injected++;
                        }
                        catch { /* log but don't crash */ }
                    }
                }

                // ── 6. Show report ────────────────────────────────────
                FamilyAuditReportWindow win =
                    new FamilyAuditReportWindow(results, stamped, injected);
                win.ShowDialog();

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // ══════════════════════════════════════════════════════════
        //  PATH RESOLUTION — Two-tier approach
        // ══════════════════════════════════════════════════════════

        /// <summary>
        /// Resolves the original .rfa path for a loaded family.
        /// 
        /// Tier 1 (fast): Read from Extensible Storage if a previous
        ///   audit or load event already stamped the family.
        /// 
        /// Tier 2 (fallback): Search the ADC workspace on disk using
        ///   Family.Name + ".rfa".  Walks the directory tree under
        ///   the user-chosen <see cref="_adcSearchRoot"/>.
        /// 
        /// CONSTRAINT: We do NOT query the ADC SQLite database.
        ///   It is locked by the Desktop Connector process and its
        ///   schema changes between ADC versions.
        /// </summary>
        /// <returns>Full path to the .rfa, or null if not found.</returns>
        private string GetAdcPathForFamily(Document doc, Family family)
        {
            // ── Tier 1: Extensible Storage ────────────────────────
            TraceData trace = FamilyTraceabilityManager.ReadTraceData(family);
            if (trace != null
                && !string.IsNullOrEmpty(trace.AdcPath)
                && File.Exists(trace.AdcPath))
            {
                return trace.AdcPath;
            }

            // ── Tier 2: Directory search ──────────────────────────
            if (string.IsNullOrEmpty(_adcSearchRoot)
                || !Directory.Exists(_adcSearchRoot))
                return null;

            string target = family.Name + ".rfa";

            try
            {
                // SearchOption.AllDirectories walks the entire ADC tree.
                // For very large workspaces this could be slow on the first
                // run; subsequent calls benefit from the Tier-1 cache once
                // families are stamped.
                string[] matches = Directory.GetFiles(
                    _adcSearchRoot, target, SearchOption.AllDirectories);

                if (matches.Length == 1)
                    return matches[0];

                if (matches.Length > 1)
                {
                    // Ambiguous — pick the one whose path contains the
                    // project name if possible, else take the first match.
                    string projName = Path.GetFileNameWithoutExtension(
                        doc.PathName ?? "");
                    string best = matches.FirstOrDefault(
                        m => m.IndexOf(projName,
                            StringComparison.OrdinalIgnoreCase) >= 0);
                    return best ?? matches[0];
                }
            }
            catch (Exception) { /* permission errors, broken symlinks, etc. */ }

            return null;
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  AUDIT REPORT WINDOW — WPF code-behind, no XAML
    // ═══════════════════════════════════════════════════════════════

    public class FamilyAuditReportWindow : Window
    {
        private readonly List<FamilyAuditItem> _items;
        private readonly int _stamped;
        private readonly int _injected;

        static readonly Color COL_BG = Color.FromRgb(245, 245, 248);
        static readonly Color COL_TEXT = Color.FromRgb(30, 30, 30);
        static readonly Color COL_SUB = Color.FromRgb(120, 120, 130);
        static readonly Color COL_ACCENT = Color.FromRgb(0, 120, 212);
        static readonly Color COL_GREEN = Color.FromRgb(40, 167, 69);
        static readonly Color COL_RED = Color.FromRgb(220, 53, 69);
        static readonly Color COL_AMBER = Color.FromRgb(230, 170, 0);
        static readonly Color COL_BORDER = Color.FromRgb(200, 200, 210);

        public FamilyAuditReportWindow(
            List<FamilyAuditItem> items, int stamped, int injected)
        {
            _items = items; _stamped = stamped; _injected = injected;
            Title = "HMV Tools - Family Audit Report";
            Width = 950; Height = 620; MinWidth = 800; MinHeight = 400;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Background = new SolidColorBrush(COL_BG);
            BuildUI();
        }

        void BuildUI()
        {
            StackPanel root = new StackPanel { Margin = new Thickness(20) };
            ScrollViewer scr = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
            };
            scr.Content = root;
            Content = scr;

            // ── Summary header ────────────────────────────────────
            int total = _items.Count;
            int match = _items.Count(i => i.Status == "Match");
            int mismatch = _items.Count(i => i.Status == "Mismatch");
            int noPath = _items.Count(i => i.Status == "No ADC Path");
            int errors = _items.Count(i => i.Status == "Error");

            root.Children.Add(new TextBlock
            {
                Text = "Family Audit Report",
                FontSize = 20,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(COL_TEXT),
                Margin = new Thickness(0, 0, 0, 4)
            });

            string summary = $"{total} families audited  ·  "
                + $"{match} match  ·  {mismatch} mismatch  ·  "
                + $"{noPath} no path  ·  {errors} errors  ·  "
                + $"{_stamped} newly stamped";
            if (_injected > 0)
                summary += $"  ·  {_injected} params injected";

            root.Children.Add(new TextBlock
            {
                Text = summary,
                FontSize = 12,
                Foreground = new SolidColorBrush(COL_SUB),
                Margin = new Thickness(0, 0, 0, 14),
                TextWrapping = TextWrapping.Wrap
            });

            // ── Column headers ────────────────────────────────────
            Grid hdr = MakeRow("Family", "Category", "Status", "ADC Path", true);
            root.Children.Add(hdr);

            // ── Result rows ───────────────────────────────────────
            var sorted = _items.OrderBy(i =>
                i.Status == "Mismatch" ? 0 :
                i.Status == "Error" ? 1 :
                i.Status == "No ADC Path" ? 2 : 3).ToList();

            foreach (FamilyAuditItem item in sorted)
            {
                Grid row = MakeRow(
                    item.FamilyName,
                    item.Category,
                    item.Status,
                    TruncPath(item.AdcPath, 60),
                    false);

                Color sc = item.Status == "Match" ? COL_GREEN :
                           item.Status == "Mismatch" ? COL_RED :
                           item.Status == "Error" ? COL_RED : COL_AMBER;

                Border bar = new Border
                {
                    Width = 4,
                    Background = new SolidColorBrush(sc),
                    VerticalAlignment = VerticalAlignment.Stretch,
                    HorizontalAlignment = HorizontalAlignment.Left
                };
                Grid.SetRowSpan(bar, 1); Grid.SetColumn(bar, 0);
                row.Children.Add(bar);

                string tip = $"Family: {item.FamilyName}\n"
                           + $"Category: {item.Category}\n"
                           + $"Status: {item.Status}\n"
                           + $"ADC Path: {item.AdcPath}\n"
                           + $"Loaded GUID: {item.LoadedGuid}\n"
                           + $"Source GUID: {item.SourceGuid}";
                if (!string.IsNullOrEmpty(item.Sha256Hash))
                    tip += $"\nSHA-256: {item.Sha256Hash}";
                if (!string.IsNullOrEmpty(item.ErrorDetail))
                    tip += $"\nError: {item.ErrorDetail}";
                row.ToolTip = tip;

                root.Children.Add(row);
            }

            // ── Action buttons ────────────────────────────────────
            StackPanel actions = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 14, 0, 0)
            };
            Button btnCopy = MkBtn("Copy to Clipboard", COL_ACCENT, Colors.White, 150);
            btnCopy.Click += (s, e) =>
            {
                var lines = new List<string>
                {
                    "Family\tCategory\tStatus\tADC Path\tLoaded GUID\tSource GUID\tSHA-256\tError"
                };
                foreach (var it in sorted)
                    lines.Add($"{it.FamilyName}\t{it.Category}\t{it.Status}\t"
                            + $"{it.AdcPath}\t{it.LoadedGuid}\t{it.SourceGuid}\t"
                            + $"{it.Sha256Hash}\t{it.ErrorDetail}");
                Clipboard.SetText(string.Join("\n", lines));
                MessageBox.Show("Copied to clipboard (tab-separated).",
                    "HMV - Family Audit", MessageBoxButton.OK,
                    MessageBoxImage.Information);
            };
            actions.Children.Add(btnCopy);

            Button btnClose = MkBtn("Close", Color.FromRgb(240, 240, 243),
                COL_TEXT, 80);
            btnClose.Click += (s, e) => Close();
            actions.Children.Add(btnClose);

            root.Children.Add(actions);
        }

        Grid MakeRow(string c1, string c2, string c3, string c4, bool isHeader)
        {
            Grid g = new Grid
            {
                Margin = new Thickness(0, 0, 0, 1),
                Background = isHeader
                    ? new SolidColorBrush(Color.FromRgb(235, 235, 240))
                    : Brushes.White
            };
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(220) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(100) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            FontWeight fw = isHeader ? FontWeights.SemiBold : FontWeights.Normal;
            double fs = 11;
            Thickness pad = new Thickness(8, 5, 4, 5);

            AddCell(g, 0, c1, fw, fs, pad, COL_TEXT);
            AddCell(g, 1, c2, fw, fs, pad, COL_SUB);
            AddCell(g, 2, c3, fw, fs, pad,
                isHeader ? COL_TEXT :
                c3 == "Match" ? COL_GREEN :
                c3 == "Mismatch" ? COL_RED :
                c3 == "Error" ? COL_RED : COL_AMBER);
            AddCell(g, 3, c4, FontWeights.Normal, 10, pad, COL_SUB);

            return g;
        }

        void AddCell(Grid g, int col, string text, FontWeight fw,
            double fs, Thickness pad, Color fg)
        {
            TextBlock tb = new TextBlock
            {
                Text = text,
                FontSize = fs,
                FontWeight = fw,
                Foreground = new SolidColorBrush(fg),
                Padding = pad,
                TextTrimming = TextTrimming.CharacterEllipsis,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(tb, col);
            g.Children.Add(tb);
        }

        string TruncPath(string p, int max)
        {
            if (string.IsNullOrEmpty(p) || p.Length <= max) return p ?? "";
            return "…" + p.Substring(p.Length - max + 1);
        }

        Button MkBtn(string text, Color bg, Color fg, double w)
        {
            Button b = new Button
            {
                Content = text,
                Width = w,
                Height = 32,
                FontSize = 11,
                Margin = new Thickness(0, 0, 6, 0),
                Cursor = Cursors.Hand,
                Foreground = new SolidColorBrush(fg),
                Background = new SolidColorBrush(bg),
                BorderThickness = new Thickness(0)
            };
            var tp = new ControlTemplate(typeof(Button));
            var bd = new FrameworkElementFactory(typeof(Border));
            bd.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
            bd.SetValue(Border.BackgroundProperty, new SolidColorBrush(bg));
            bd.SetValue(Border.PaddingProperty, new Thickness(10, 4, 10, 4));
            var cp = new FrameworkElementFactory(typeof(ContentPresenter));
            cp.SetValue(ContentPresenter.HorizontalAlignmentProperty,
                HorizontalAlignment.Center);
            cp.SetValue(ContentPresenter.VerticalAlignmentProperty,
                VerticalAlignment.Center);
            bd.AppendChild(cp);
            tp.VisualTree = bd;
            b.Template = tp;
            return b;
        }
    }
}