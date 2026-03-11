using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

// Resolve Revit UI vs WPF type conflicts
using Color = System.Windows.Media.Color;
using TextBox = System.Windows.Controls.TextBox;
using ComboBox = System.Windows.Controls.ComboBox;
using CheckBox = System.Windows.Controls.CheckBox;
using Button = System.Windows.Controls.Button;
using Grid = System.Windows.Controls.Grid;
using Line = System.Windows.Shapes.Line;
using Ellipse = System.Windows.Shapes.Ellipse;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class USDTuberiaConfigCommand : IExternalCommand
    {
        public const int MAX_ROWS = 5;
        public const int MAX_COLS = 9;
        public const int MAX_DIST = 9;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // ── Get selected family instances ──
                ICollection<ElementId> selIds = uidoc.Selection.GetElementIds();
                List<FamilyInstance> instances = new List<FamilyInstance>();

                foreach (ElementId id in selIds)
                {
                    Element el = doc.GetElement(id);
                    if (el is FamilyInstance fi)
                        instances.Add(fi);
                }

                if (instances.Count == 0)
                {
                    TaskDialog.Show("HMV - Ductos Editor",
                        "Seleccione al menos una instancia de la familia USD\n" +
                        "antes de ejecutar el comando.");
                    return Result.Cancelled;
                }

                FamilyInstance primary = instances[0];

                // ══════════════════════════════════════════
                //  STEP 1: Matrix size selector
                // ══════════════════════════════════════════
                USDMatrixSizeWindow sizeWin = new USDMatrixSizeWindow();
                if (sizeWin.ShowDialog() != true)
                    return Result.Cancelled;

                int rows = sizeWin.SelectedRows;
                int cols = sizeWin.SelectedCols;

                // ══════════════════════════════════════════
                //  STEP 2: Read current parameters
                // ══════════════════════════════════════════
                string diaPrefix = FindDiameterPrefix(primary);
                USDTuberiaData data = ReadParameters(primary, diaPrefix);

                if (data == null)
                    return Result.Failed;

                // ══════════════════════════════════════════
                //  STEP 3: Main editor
                // ══════════════════════════════════════════
                var typeGroups = instances
                    .GroupBy(fi => fi.Symbol.Id).ToList();
                string typeInfo = $"Tipo actual: {primary.Symbol.Name}";

                USDTuberiaWindow win = new USDTuberiaWindow(
                    data, instances.Count, typeInfo, rows, cols);

                if (win.ShowDialog() != true)
                    return Result.Cancelled;

                USDTuberiaData newData = win.ResultData;

                // Turn off all holes outside the matrix
                for (int r = 0; r < MAX_ROWS; r++)
                    for (int c = 0; c < MAX_COLS; c++)
                        if (r >= rows || c >= cols)
                            newData.Visible[r, c] = false;

                // ══════════════════════════════════════════
                //  STEP 4: Build type name & find/create
                // ══════════════════════════════════════════
                string targetName = BuildTypeName(newData, rows, cols);

                using (Transaction tx = new Transaction(doc,
                    "HMV - Configurar USD Tuberías"))
                {
                    tx.Start();

                    FamilySymbol targetType = FindOrCreateType(
                        doc, primary, targetName, newData, diaPrefix);

                    if (!targetType.IsActive)
                        targetType.Activate();

                    // Write parameters to the target type
                    WriteParametersToType(targetType, newData, diaPrefix);

                    // Change all selected instances to this type
                    foreach (FamilyInstance fi in instances)
                    {
                        if (fi.Symbol.Id != targetType.Id)
                            fi.ChangeTypeId(targetType.Id);
                    }

                    tx.Commit();
                }

                // ══════════════════════════════════════════
                //  STEP 5: Optionally purge unused types
                // ══════════════════════════════════════════
                PurgeUnusedTypes(doc, primary.Symbol.Family);

                TaskDialog.Show("HMV - Ductos Editor",
                    $"Configuración aplicada.\n\n" +
                    $"Tipo: {targetName}\n" +
                    $"Instancias actualizadas: {instances.Count}\n" +
                    $"Matriz: {rows} × {cols}");

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

        // ═══════════════════════════════════════════════════════
        //  Type name builder
        // ═══════════════════════════════════════════════════════

        private string BuildTypeName(USDTuberiaData data,
            int rows, int cols)
        {
            // Count visible holes
            int holeCount = 0;
            List<double> diameters = new List<double>();

            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    if (data.Visible[r, c])
                    {
                        holeCount++;
                        double d = data.Diameter[r, c];
                        if (d > 0 && !diameters.Any(
                            x => Math.Abs(x - d) < 0.1))
                            diameters.Add(d);
                    }

            diameters.Sort();

            // Format diameters in inches
            string diaStr;
            if (diameters.Count == 0)
                diaStr = "0\"";
            else if (diameters.Count == 1)
                diaStr = FormatInches(diameters[0]);
            else
                diaStr = "(" + string.Join(",",
                    diameters.Select(d => FormatInches(d))) + ")";

            // Width × Height in meters
            double wM = data.Ancho / 1000.0;
            double hM = data.Altura / 1000.0;

            return $"{holeCount} \u00f8 {diaStr} ({wM:F2} m x {hM:F2} m)";
        }

        private string FormatInches(double mm)
        {
            double inches = mm / 25.4;
            // Check if it's a clean fraction
            if (Math.Abs(inches - Math.Round(inches)) < 0.05)
                return $"{Math.Round(inches):F0}\"";
            else
                return $"{inches:F1}\"";
        }

        // ═══════════════════════════════════════════════════════
        //  Type management
        // ═══════════════════════════════════════════════════════

        private FamilySymbol FindOrCreateType(Document doc,
            FamilyInstance source, string targetName,
            USDTuberiaData data, string diaPrefix)
        {
            Family family = source.Symbol.Family;

            // Search existing types in this family
            foreach (ElementId typeId in family.GetFamilySymbolIds())
            {
                FamilySymbol fs = doc.GetElement(typeId) as FamilySymbol;
                if (fs != null && fs.Name == targetName)
                    return fs;
            }

            // Not found — duplicate from current type
            ElementId newTypeId = source.Symbol.Duplicate(targetName).Id;
            return doc.GetElement(newTypeId) as FamilySymbol;
        }

        private void PurgeUnusedTypes(Document doc, Family family)
        {
            // Collect types with 0 instances
            List<FamilySymbol> unused = new List<FamilySymbol>();
            ICollection<ElementId> allTypeIds = family.GetFamilySymbolIds();

            // Don't purge if only 1 type exists
            if (allTypeIds.Count <= 1) return;

            foreach (ElementId typeId in allTypeIds)
            {
                FamilySymbol fs = doc.GetElement(typeId) as FamilySymbol;
                if (fs == null) continue;

                // Count instances of this type
                int count = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Count(fi => fi.Symbol.Id == typeId);

                if (count == 0)
                    unused.Add(fs);
            }

            if (unused.Count == 0) return;

            // Confirm before deleting
            string names = string.Join("\n",
                unused.Select(u => "  • " + u.Name));

            MessageBoxResult result = MessageBox.Show(
                $"Se encontraron {unused.Count} tipo(s) sin instancias:\n\n" +
                names +
                "\n\n¿Desea eliminarlos?",
                "HMV - Purgar tipos sin uso",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes) return;

            using (Transaction tx = new Transaction(doc,
                "HMV - Purgar tipos USD sin uso"))
            {
                tx.Start();
                foreach (FamilySymbol fs in unused)
                {
                    try { doc.Delete(fs.Id); }
                    catch { /* skip if can't delete */ }
                }
                tx.Commit();
            }
        }

        // ═══════════════════════════════════════════════════════
        //  Parameter read / write
        // ═══════════════════════════════════════════════════════

        private string FindDiameterPrefix(FamilyInstance fi)
        {
            FamilySymbol sym = fi.Symbol;
            string[] candidates = new string[]
            {
                "HMV_USD_\u00f8 Tuberia ",
                "HMV_USD_\u00d8 Tuberia ",
                "HMV_USD_o Tuberia ",
                "HMV_USD_O Tuberia ",
                "HMV_USD_\u00f8 TUBERIA ",
                "HMV_USD_\u00d8 TUBERIA ",
                "HMV_USD_\u2300 Tuberia ",
            };

            foreach (string prefix in candidates)
            {
                if (sym.LookupParameter(prefix + "1-1") != null)
                    return prefix;
            }

            // Scan all type parameters
            foreach (Parameter p in sym.Parameters)
            {
                string pName = p.Definition?.Name ?? "";
                if (pName.Contains("Tuberia 1-1") &&
                    !pName.Contains("Visibilidad") &&
                    !pName.Contains("Distancia"))
                {
                    int idx = pName.IndexOf("1-1");
                    if (idx > 0) return pName.Substring(0, idx);
                }
            }

            return "HMV_USD_\u00f8 Tuberia ";
        }

        private USDTuberiaData ReadParameters(FamilyInstance fi,
            string diaPrefix)
        {
            USDTuberiaData data = new USDTuberiaData();
            List<string> missing = new List<string>();

            data.Ancho = TryGetDouble(fi, "HMV_USD_Ancho", ref missing);
            data.Altura = TryGetDouble(fi, "HMV_USD_Altura", ref missing);

            for (int i = 1; i <= MAX_DIST; i++)
            {
                data.DistX[i - 1] = TryGetDouble(fi,
                    $"HMV_USD_Distancia {i}_X", ref missing);
                data.DistY[i - 1] = TryGetDouble(fi,
                    $"HMV_USD_Distancia {i}_Y", ref missing);
            }

            for (int r = 1; r <= MAX_ROWS; r++)
                for (int c = 1; c <= MAX_COLS; c++)
                {
                    data.Visible[r - 1, c - 1] = TryGetBool(fi,
                        $"HMV_USD_Visibilidad Tuberia {r}-{c}", ref missing);
                    data.Diameter[r - 1, c - 1] = TryGetDouble(fi,
                        $"{diaPrefix}{r}-{c}", ref missing);
                }

            if (missing.Count > 20)
            {
                string allP = GetAllParameterNames(fi);
                TaskDialog td = new TaskDialog("HMV - Diagnóstico");
                td.MainInstruction = $"{missing.Count} parámetros faltantes.";
                td.MainContent =
                    "Primeros no encontrados:\n" +
                    string.Join("\n", missing.Take(8).Select(m => "  • " + m)) +
                    "\n\nDisponibles:\n" + allP;
                td.Show();
                return null;
            }
            else if (missing.Count > 0)
            {
                TaskDialog td = new TaskDialog("HMV - Aviso");
                td.MainInstruction = $"{missing.Count} parámetro(s) no encontrado(s).";
                td.MainContent = string.Join("\n",
                    missing.Take(10).Select(m => "  • " + m));
                td.Show();
            }

            return data;
        }

        private void WriteParametersToType(FamilySymbol sym,
            USDTuberiaData data, string diaPrefix)
        {
            SetDoubleOnType(sym, "HMV_USD_Ancho", data.Ancho);
            SetDoubleOnType(sym, "HMV_USD_Altura", data.Altura);

            for (int i = 1; i <= MAX_DIST; i++)
            {
                SetDoubleOnType(sym,
                    $"HMV_USD_Distancia {i}_X", data.DistX[i - 1]);
                SetDoubleOnType(sym,
                    $"HMV_USD_Distancia {i}_Y", data.DistY[i - 1]);
            }

            for (int r = 1; r <= MAX_ROWS; r++)
                for (int c = 1; c <= MAX_COLS; c++)
                {
                    SetBoolOnType(sym,
                        $"HMV_USD_Visibilidad Tuberia {r}-{c}",
                        data.Visible[r - 1, c - 1]);
                    SetDoubleOnType(sym,
                        $"{diaPrefix}{r}-{c}",
                        data.Diameter[r - 1, c - 1]);
                }
        }

        // ── Helpers ──

        private string GetAllParameterNames(FamilyInstance fi)
        {
            FamilySymbol sym = fi.Symbol;
            List<string> names = new List<string>();
            foreach (Parameter p in sym.Parameters)
            {
                string name = p.Definition?.Name ?? "(null)";
                if (name.StartsWith("HMV"))
                    names.Add("[Type] " + name);
            }
            names.Sort();
            return string.Join("\n",
                names.Take(30).Select(n => "  " + n));
        }

        private double TryGetDouble(FamilyInstance fi, string name,
            ref List<string> missing)
        {
            Parameter p = fi.Symbol.LookupParameter(name);
            if (p == null) { missing.Add(name); return 0; }
            return p.AsDouble() * 304.8;
        }

        private bool TryGetBool(FamilyInstance fi, string name,
            ref List<string> missing)
        {
            Parameter p = fi.Symbol.LookupParameter(name);
            if (p == null) { missing.Add(name); return false; }
            return p.AsInteger() == 1;
        }

        private void SetDoubleOnType(FamilySymbol sym,
            string name, double mm)
        {
            Parameter p = sym.LookupParameter(name);
            if (p != null && !p.IsReadOnly)
                p.Set(mm / 304.8);
        }

        private void SetBoolOnType(FamilySymbol sym,
            string name, bool val)
        {
            Parameter p = sym.LookupParameter(name);
            if (p != null && !p.IsReadOnly)
                p.Set(val ? 1 : 0);
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  Data model
    // ═══════════════════════════════════════════════════════════════

    public class USDTuberiaData
    {
        public double Ancho;
        public double Altura;
        public double[] DistX = new double[USDTuberiaConfigCommand.MAX_DIST];
        public double[] DistY = new double[USDTuberiaConfigCommand.MAX_DIST];
        public bool[,] Visible =
            new bool[USDTuberiaConfigCommand.MAX_ROWS,
                      USDTuberiaConfigCommand.MAX_COLS];
        public double[,] Diameter =
            new double[USDTuberiaConfigCommand.MAX_ROWS,
                        USDTuberiaConfigCommand.MAX_COLS];

        public USDTuberiaData Clone()
        {
            USDTuberiaData c = new USDTuberiaData();
            c.Ancho = Ancho;
            c.Altura = Altura;
            Array.Copy(DistX, c.DistX, DistX.Length);
            Array.Copy(DistY, c.DistY, DistY.Length);
            Array.Copy(Visible, c.Visible, Visible.Length);
            Array.Copy(Diameter, c.Diameter, Diameter.Length);
            return c;
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  STEP 1 WINDOW: Matrix size selector
    // ═══════════════════════════════════════════════════════════════

    public class USDMatrixSizeWindow : Window
    {
        public int SelectedRows { get; private set; } = 5;
        public int SelectedCols { get; private set; } = 9;

        private ComboBox _cmbRows;
        private ComboBox _cmbCols;
        private Canvas _previewCanvas;

        private static readonly Color COL_BG = Color.FromRgb(245, 245, 248);
        private static readonly Color COL_ACCENT = Color.FromRgb(0, 120, 212);
        private static readonly Color COL_TEXT = Color.FromRgb(30, 30, 30);
        private static readonly Color COL_SUB = Color.FromRgb(120, 120, 130);
        private static readonly Color COL_BORDER = Color.FromRgb(200, 200, 210);
        private static readonly Color COL_GREEN = Color.FromRgb(40, 167, 69);

        public USDMatrixSizeWindow()
        {
            Title = "HMV Tools - Ductos Editor";
            Width = 480;
            Height = 520;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(COL_BG);

            StackPanel root = new StackPanel
            {
                Margin = new Thickness(24)
            };
            Content = root;

            // Title
            root.Children.Add(new TextBlock
            {
                Text = "¿De cuánto por cuánto es la\nmatriz de ductos?",
                FontSize = 20,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(COL_TEXT),
                Margin = new Thickness(0, 0, 0, 6),
                TextWrapping = TextWrapping.Wrap
            });

            root.Children.Add(new TextBlock
            {
                Text = "Seleccione el número de filas y columnas de tuberías.",
                FontSize = 12,
                Foreground = new SolidColorBrush(COL_SUB),
                Margin = new Thickness(0, 0, 0, 20)
            });

            // Row/Col selectors
            Grid selGrid = new Grid();
            selGrid.ColumnDefinitions.Add(
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            selGrid.ColumnDefinitions.Add(
                new ColumnDefinition { Width = new GridLength(30) });
            selGrid.ColumnDefinitions.Add(
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            selGrid.Margin = new Thickness(0, 0, 0, 20);
            root.Children.Add(selGrid);

            // Rows
            StackPanel rowPanel = new StackPanel();
            rowPanel.Children.Add(new TextBlock
            {
                Text = "Filas (Y)",
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(COL_ACCENT),
                Margin = new Thickness(0, 0, 0, 6)
            });
            _cmbRows = new ComboBox
            {
                Height = 36,
                FontSize = 14,
                VerticalContentAlignment = VerticalAlignment.Center
            };
            for (int i = 1; i <= 5; i++) _cmbRows.Items.Add(i);
            _cmbRows.SelectedIndex = 4; // default 5
            _cmbRows.SelectionChanged += (s, e) => DrawPreview();
            rowPanel.Children.Add(_cmbRows);
            Grid.SetColumn(rowPanel, 0);
            selGrid.Children.Add(rowPanel);

            // "×" label
            TextBlock xLabel = new TextBlock
            {
                Text = "×",
                FontSize = 22,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(COL_TEXT),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Bottom,
                Margin = new Thickness(0, 0, 0, 6)
            };
            Grid.SetColumn(xLabel, 1);
            selGrid.Children.Add(xLabel);

            // Cols
            StackPanel colPanel = new StackPanel();
            colPanel.Children.Add(new TextBlock
            {
                Text = "Columnas (X)",
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(COL_ACCENT),
                Margin = new Thickness(0, 0, 0, 6)
            });
            _cmbCols = new ComboBox
            {
                Height = 36,
                FontSize = 14,
                VerticalContentAlignment = VerticalAlignment.Center
            };
            for (int i = 1; i <= 9; i++) _cmbCols.Items.Add(i);
            _cmbCols.SelectedIndex = 8; // default 9
            _cmbCols.SelectionChanged += (s, e) => DrawPreview();
            colPanel.Children.Add(_cmbCols);
            Grid.SetColumn(colPanel, 2);
            selGrid.Children.Add(colPanel);

            // Preview
            Border previewBorder = new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(COL_BORDER),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Height = 160,
                Margin = new Thickness(0, 0, 0, 20)
            };
            _previewCanvas = new Canvas
            {
                ClipToBounds = true
            };
            _previewCanvas.Loaded += (s, e) => DrawPreview();
            previewBorder.Child = _previewCanvas;
            root.Children.Add(previewBorder);

            // Buttons
            StackPanel btnRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            Button cancelBtn = MakeBtn("Cancelar",
                Color.FromRgb(240, 240, 243), COL_TEXT, 100);
            cancelBtn.Click += (s, e) =>
            {
                DialogResult = false;
                Close();
            };
            btnRow.Children.Add(cancelBtn);

            Button okBtn = MakeBtn("Continuar →",
                COL_ACCENT, Colors.White, 140);
            okBtn.FontWeight = FontWeights.SemiBold;
            okBtn.Click += (s, e) =>
            {
                SelectedRows = (int)_cmbRows.SelectedItem;
                SelectedCols = (int)_cmbCols.SelectedItem;
                DialogResult = true;
                Close();
            };
            btnRow.Children.Add(okBtn);
            root.Children.Add(btnRow);
        }

        private void DrawPreview()
        {
            _previewCanvas.Children.Clear();
            double w = _previewCanvas.ActualWidth;
            double h = _previewCanvas.ActualHeight;
            if (w < 10 || h < 10) return;

            int r = (int)(_cmbRows.SelectedItem ?? 5);
            int c = (int)(_cmbCols.SelectedItem ?? 9);

            double mL = 30, mT = 20;
            double cW = (w - mL - 20) / Math.Max(c, 1);
            double cH = (h - mT - 20) / Math.Max(r, 1);
            double rad = Math.Min(Math.Min(cW, cH) * 0.35, 12);

            // Label
            TextBlock lbl = new TextBlock
            {
                Text = $"Matriz {r} × {c}  =  {r * c} posiciones",
                FontSize = 10,
                Foreground = new SolidColorBrush(COL_SUB)
            };
            Canvas.SetLeft(lbl, 4);
            Canvas.SetTop(lbl, 2);
            _previewCanvas.Children.Add(lbl);

            for (int ri = 0; ri < r; ri++)
                for (int ci = 0; ci < c; ci++)
                {
                    double cx = mL + ci * cW + cW / 2;
                    double cy = mT + ri * cH + cH / 2;

                    Ellipse el = new Ellipse
                    {
                        Width = rad * 2,
                        Height = rad * 2,
                        Stroke = new SolidColorBrush(COL_GREEN),
                        StrokeThickness = 1.5,
                        Fill = new SolidColorBrush(
                            Color.FromArgb(80, 40, 167, 69))
                    };
                    Canvas.SetLeft(el, cx - rad);
                    Canvas.SetTop(el, cy - rad);
                    _previewCanvas.Children.Add(el);
                }
        }

        private Button MakeBtn(string text, Color bg,
            Color fg, double width)
        {
            Button btn = new Button
            {
                Content = text,
                Width = width,
                Height = 36,
                FontSize = 13,
                Margin = new Thickness(0, 0, 8, 0),
                Cursor = Cursors.Hand,
                Foreground = new SolidColorBrush(fg),
                Background = new SolidColorBrush(bg),
                BorderThickness = new Thickness(0)
            };
            ControlTemplate t = new ControlTemplate(typeof(Button));
            FrameworkElementFactory bdr =
                new FrameworkElementFactory(typeof(Border));
            bdr.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
            bdr.SetValue(Border.BackgroundProperty, new SolidColorBrush(bg));
            bdr.SetValue(Border.PaddingProperty, new Thickness(14, 6, 14, 6));
            FrameworkElementFactory cp =
                new FrameworkElementFactory(typeof(ContentPresenter));
            cp.SetValue(ContentPresenter.HorizontalAlignmentProperty,
                HorizontalAlignment.Center);
            cp.SetValue(ContentPresenter.VerticalAlignmentProperty,
                VerticalAlignment.Center);
            bdr.AppendChild(cp);
            t.VisualTree = bdr;
            btn.Template = t;
            return btn;
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  STEP 2 WINDOW: Main editor (dynamic grid)
    // ═══════════════════════════════════════════════════════════════

    public class USDTuberiaWindow : Window
    {
        public USDTuberiaData ResultData { get; private set; }

        private readonly USDTuberiaData _original;
        private readonly int _instanceCount;
        private readonly string _typeInfo;
        private readonly int _rows;
        private readonly int _cols;

        private TextBox _txtAncho;
        private TextBox _txtAltura;
        private TextBox[] _txtDistX;
        private TextBox[] _txtDistY;
        private CheckBox[,] _chkVisible;
        private TextBox[,] _txtDiameter;
        private Border[,] _cellBorders;
        private Canvas _schematicCanvas;

        // ── HMV palette ──
        private static readonly Color COL_BG = Color.FromRgb(245, 245, 248);
        private static readonly Color COL_ACCENT = Color.FromRgb(0, 120, 212);
        private static readonly Color COL_GREEN = Color.FromRgb(40, 167, 69);
        private static readonly Color COL_RED = Color.FromRgb(220, 53, 69);
        private static readonly Color COL_BORDER = Color.FromRgb(200, 200, 210);
        private static readonly Color COL_CELL_ON = Color.FromRgb(232, 245, 233);
        private static readonly Color COL_CELL_OFF = Color.FromRgb(245, 245, 245);
        private static readonly Color COL_TEXT = Color.FromRgb(30, 30, 30);
        private static readonly Color COL_SUB = Color.FromRgb(120, 120, 130);
        private static readonly Color COL_BTN_BG = Color.FromRgb(240, 240, 243);

        public USDTuberiaWindow(USDTuberiaData data, int instanceCount,
            string typeInfo, int rows, int cols)
        {
            _original = data;
            _instanceCount = instanceCount;
            _typeInfo = typeInfo;
            _rows = rows;
            _cols = cols;
            ResultData = data.Clone();

            _txtDistX = new TextBox[_cols];
            _txtDistY = new TextBox[_rows];
            _chkVisible = new CheckBox[_rows, _cols];
            _txtDiameter = new TextBox[_rows, _cols];
            _cellBorders = new Border[_rows, _cols];

            BuildUI();
        }

        private void BuildUI()
        {
            Title = "HMV Tools - Ductos Editor";
            Width = Math.Max(700, 200 + _cols * 105);
            Height = Math.Max(600, 400 + _rows * 60);
            MinWidth = 700;
            MinHeight = 600;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Background = new SolidColorBrush(COL_BG);
            ResizeMode = ResizeMode.CanResize;

            ScrollViewer scroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                Padding = new Thickness(24)
            };
            Content = scroll;

            StackPanel root = new StackPanel();
            scroll.Content = root;

            // ── HEADER ──
            root.Children.Add(new TextBlock
            {
                Text = $"Ductos Editor  —  Matriz {_rows} × {_cols}",
                FontSize = 22,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(COL_TEXT),
                Margin = new Thickness(0, 0, 0, 2)
            });

            root.Children.Add(new TextBlock
            {
                Text = $"{_instanceCount} instancia(s)  ·  {_typeInfo}",
                FontSize = 12,
                Foreground = new SolidColorBrush(COL_SUB),
                Margin = new Thickness(0, 0, 0, 16),
                TextWrapping = TextWrapping.Wrap
            });

            // ── TOP: Schematic + Dimensions/Distances ──
            Grid topRow = new Grid();
            topRow.ColumnDefinitions.Add(
                new ColumnDefinition { Width = new GridLength(320) });
            topRow.ColumnDefinitions.Add(
                new ColumnDefinition
                {
                    Width = new GridLength(
                    1, GridUnitType.Star)
                });
            topRow.Margin = new Thickness(0, 0, 0, 16);
            root.Children.Add(topRow);

            // Schematic
            Border schemBorder = MakeCard();
            schemBorder.MinHeight = 200;
            schemBorder.Margin = new Thickness(0, 0, 12, 0);
            _schematicCanvas = new Canvas
            {
                Background = Brushes.White,
                ClipToBounds = true
            };
            _schematicCanvas.Loaded += (s, e) => DrawSchematic();
            schemBorder.Child = _schematicCanvas;
            Grid.SetColumn(schemBorder, 0);
            topRow.Children.Add(schemBorder);

            // Distances + dimensions panel
            Border distCard = MakeCard();
            distCard.Padding = new Thickness(16);
            StackPanel distPanel = new StackPanel();
            distCard.Child = distPanel;
            Grid.SetColumn(distCard, 1);
            topRow.Children.Add(distCard);

            // Framing dimensions
            distPanel.Children.Add(SectionHeader("Dimensiones del marco (mm)"));
            WrapPanel wpDims = new WrapPanel
            {
                Margin = new Thickness(0, 4, 0, 12)
            };
            wpDims.Children.Add(MakeDistField("Ancho",
                out _txtAncho, ResultData.Ancho));
            wpDims.Children.Add(MakeDistField("Altura",
                out _txtAltura, ResultData.Altura));
            distPanel.Children.Add(wpDims);

            // Distance X (only _cols fields)
            distPanel.Children.Add(SectionHeader("Distancias (mm)"));
            distPanel.Children.Add(SubLabel(
                $"Distancia X — entre columnas (X1..X{_cols})"));
            WrapPanel wpX = new WrapPanel
            {
                Margin = new Thickness(0, 4, 0, 10)
            };
            for (int i = 0; i < _cols; i++)
                wpX.Children.Add(MakeDistField($"X{i + 1}",
                    out _txtDistX[i], ResultData.DistX[i]));
            distPanel.Children.Add(wpX);

            // Distance Y (only _rows fields)
            distPanel.Children.Add(SubLabel(
                $"Distancia Y — entre filas (Y1..Y{_rows})"));
            WrapPanel wpY = new WrapPanel
            {
                Margin = new Thickness(0, 4, 0, 10)
            };
            for (int i = 0; i < _rows; i++)
                wpY.Children.Add(MakeDistField($"Y{i + 1}",
                    out _txtDistY[i], ResultData.DistY[i]));
            distPanel.Children.Add(wpY);

            distPanel.Children.Add(BuildDistQuickApply());

            // ── GRID ──
            root.Children.Add(SectionHeader(
                $"Tuberías — Visibilidad y Diámetro (mm)  " +
                $"[{_rows}×{_cols}]"));

            Border gridCard = MakeCard();
            gridCard.Padding = new Thickness(12);
            gridCard.Margin = new Thickness(0, 6, 0, 12);
            gridCard.Child = BuildTuberiaGrid();
            root.Children.Add(gridCard);

            // ── BULK TOOLS ──
            root.Children.Add(SectionHeader("Herramientas masivas"));
            Border bulkCard = MakeCard();
            bulkCard.Padding = new Thickness(12);
            bulkCard.Margin = new Thickness(0, 6, 0, 16);
            bulkCard.Child = BuildBulkTools();
            root.Children.Add(bulkCard);

            // ── ACTIONS ──
            root.Children.Add(BuildActionButtons());
        }

        // ═══════════════════════════════════════════════════════
        //  Schematic
        // ═══════════════════════════════════════════════════════

        private void DrawSchematic()
        {
            _schematicCanvas.Children.Clear();
            double w = _schematicCanvas.ActualWidth;
            double h = _schematicCanvas.ActualHeight;
            if (w < 10 || h < 10) return;

            double mL = 40, mT = 36;
            double cW = (w - mL - 20) / _cols;
            double cH = (h - mT - 30) / _rows;
            double rad = Math.Min(Math.Min(cW, cH) * 0.35, 14);

            AddCanvasText("Esquema de referencia", 4, 4, 9,
                FontWeights.SemiBold, COL_SUB);

            // Baselines
            AddDashLine(mL - 10, mT - 8, mL - 10,
                mT + _rows * cH + 5, COL_RED);
            AddDashLine(mL - 10, mT - 8,
                mL + _cols * cW + 5, mT - 8, COL_RED);
            AddCanvasText("0,0", mL - 28, mT - 24, 8,
                FontWeights.Bold, COL_RED);

            for (int c = 0; c < _cols; c++)
                AddCanvasText($"X{c + 1}",
                    mL + c * cW + cW / 2 - 8, mT - 22, 7.5,
                    FontWeights.Normal, COL_ACCENT);
            for (int r = 0; r < _rows; r++)
                AddCanvasText($"Y{r + 1}",
                    4, mT + r * cH + cH / 2 - 7, 7.5,
                    FontWeights.Normal, COL_ACCENT);

            for (int r = 0; r < _rows; r++)
                for (int c = 0; c < _cols; c++)
                {
                    double cx = mL + c * cW + cW / 2;
                    double cy = mT + r * cH + cH / 2;
                    bool vis = ResultData.Visible[r, c];

                    Ellipse el = new Ellipse
                    {
                        Width = rad * 2,
                        Height = rad * 2,
                        Stroke = new SolidColorBrush(
                            vis ? COL_GREEN
                                : Color.FromRgb(200, 200, 200)),
                        StrokeThickness = vis ? 1.8 : 1,
                        Fill = vis
                            ? new SolidColorBrush(
                                Color.FromArgb(100, 40, 167, 69))
                            : Brushes.Transparent
                    };
                    Canvas.SetLeft(el, cx - rad);
                    Canvas.SetTop(el, cy - rad);
                    _schematicCanvas.Children.Add(el);

                    AddCanvasText($"{r + 1}-{c + 1}",
                        cx - 8, cy - 6, 7,
                        FontWeights.Normal,
                        vis ? Colors.White
                            : Color.FromRgb(180, 180, 180));
                }
        }

        private void AddCanvasText(string text, double x, double y,
            double size, FontWeight weight, Color color)
        {
            TextBlock tb = new TextBlock
            {
                Text = text,
                FontSize = size,
                FontWeight = weight,
                Foreground = new SolidColorBrush(color)
            };
            Canvas.SetLeft(tb, x);
            Canvas.SetTop(tb, y);
            _schematicCanvas.Children.Add(tb);
        }

        private void AddDashLine(double x1, double y1,
            double x2, double y2, Color color)
        {
            Line ln = new Line
            {
                X1 = x1,
                Y1 = y1,
                X2 = x2,
                Y2 = y2,
                Stroke = new SolidColorBrush(color),
                StrokeThickness = 1.5,
                StrokeDashArray = new DoubleCollection { 4, 3 }
            };
            _schematicCanvas.Children.Add(ln);
        }

        // ═══════════════════════════════════════════════════════
        //  Distance fields
        // ═══════════════════════════════════════════════════════

        private StackPanel MakeDistField(string label,
            out TextBox txt, double value)
        {
            StackPanel sp = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 8, 4)
            };
            sp.Children.Add(new TextBlock
            {
                Text = label + ":",
                FontSize = 11,
                Foreground = new SolidColorBrush(COL_TEXT),
                VerticalAlignment = VerticalAlignment.Center,
                Width = 42
            });
            Border bdr = MakeInputBorder();
            txt = new TextBox
            {
                Width = 56,
                Height = 26,
                FontSize = 11,
                Text = value.ToString("F1"),
                TextAlignment = TextAlignment.Center,
                VerticalContentAlignment = VerticalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(2, 0, 2, 0)
            };
            WireFocus(txt, bdr);
            bdr.Child = txt;
            sp.Children.Add(bdr);
            return sp;
        }

        private UIElement BuildDistQuickApply()
        {
            StackPanel row = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 6, 0, 0)
            };

            row.Children.Add(SmallLabel("Todos X:"));
            TextBox txX = SmallInput(52);
            Border bx = MakeInputBorder();
            WireFocus(txX, bx); bx.Child = txX;
            row.Children.Add(bx);
            Button btnX = SmallBtn("Aplicar");
            btnX.Click += (s, e) =>
            {
                if (double.TryParse(txX.Text, out double v))
                    foreach (TextBox t in _txtDistX) t.Text = v.ToString("F1");
            };
            row.Children.Add(btnX);

            row.Children.Add(new Border { Width = 16 });

            row.Children.Add(SmallLabel("Todos Y:"));
            TextBox txY = SmallInput(52);
            Border by = MakeInputBorder();
            WireFocus(txY, by); by.Child = txY;
            row.Children.Add(by);
            Button btnY = SmallBtn("Aplicar");
            btnY.Click += (s, e) =>
            {
                if (double.TryParse(txY.Text, out double v))
                    foreach (TextBox t in _txtDistY) t.Text = v.ToString("F1");
            };
            row.Children.Add(btnY);

            return row;
        }

        // ═══════════════════════════════════════════════════════
        //  Dynamic grid (_rows × _cols)
        // ═══════════════════════════════════════════════════════

        private Grid BuildTuberiaGrid()
        {
            Grid grid = new Grid();
            grid.ColumnDefinitions.Add(
                new ColumnDefinition { Width = new GridLength(56) });
            for (int c = 0; c < _cols; c++)
                grid.ColumnDefinitions.Add(
                    new ColumnDefinition
                    {
                        Width = new GridLength(
                        1, GridUnitType.Star)
                    });

            grid.RowDefinitions.Add(
                new RowDefinition { Height = GridLength.Auto });
            for (int r = 0; r < _rows; r++)
                grid.RowDefinitions.Add(
                    new RowDefinition { Height = GridLength.Auto });

            for (int c = 0; c < _cols; c++)
            {
                TextBlock hdr = new TextBlock
                {
                    Text = $"Col {c + 1}",
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush(COL_ACCENT),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 0, 0, 6)
                };
                Grid.SetRow(hdr, 0);
                Grid.SetColumn(hdr, c + 1);
                grid.Children.Add(hdr);
            }

            for (int r = 0; r < _rows; r++)
            {
                TextBlock rl = new TextBlock
                {
                    Text = $"Fila {r + 1}",
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush(
                        Color.FromRgb(60, 60, 60)),
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Margin = new Thickness(0, 0, 8, 0)
                };
                Grid.SetRow(rl, r + 1);
                Grid.SetColumn(rl, 0);
                grid.Children.Add(rl);

                for (int c = 0; c < _cols; c++)
                {
                    int row = r, col = c;

                    Border cell = new Border
                    {
                        CornerRadius = new CornerRadius(6),
                        BorderBrush = new SolidColorBrush(COL_BORDER),
                        BorderThickness = new Thickness(1),
                        Background = new SolidColorBrush(
                            ResultData.Visible[r, c]
                                ? COL_CELL_ON : COL_CELL_OFF),
                        Margin = new Thickness(2),
                        Padding = new Thickness(4),
                        MinWidth = 80
                    };
                    _cellBorders[r, c] = cell;

                    StackPanel cs = new StackPanel();

                    CheckBox chk = new CheckBox
                    {
                        Content = $"{r + 1}-{c + 1}",
                        FontSize = 10,
                        IsChecked = ResultData.Visible[r, c],
                        Foreground = new SolidColorBrush(COL_TEXT),
                        Margin = new Thickness(0, 0, 0, 3)
                    };
                    chk.Checked += (s, e) =>
                    {
                        _cellBorders[row, col].Background =
                            new SolidColorBrush(COL_CELL_ON);
                        _txtDiameter[row, col].IsEnabled = true;
                    };
                    chk.Unchecked += (s, e) =>
                    {
                        _cellBorders[row, col].Background =
                            new SolidColorBrush(COL_CELL_OFF);
                        _txtDiameter[row, col].IsEnabled = false;
                    };
                    _chkVisible[r, c] = chk;
                    cs.Children.Add(chk);

                    StackPanel dr = new StackPanel
                    {
                        Orientation = Orientation.Horizontal
                    };
                    TextBox txt = new TextBox
                    {
                        Width = 48,
                        Height = 22,
                        FontSize = 10,
                        Text = ResultData.Diameter[r, c].ToString("F1"),
                        TextAlignment = TextAlignment.Center,
                        VerticalContentAlignment = VerticalAlignment.Center,
                        BorderThickness = new Thickness(1),
                        BorderBrush = new SolidColorBrush(COL_BORDER),
                        IsEnabled = ResultData.Visible[r, c],
                        Padding = new Thickness(2, 0, 2, 0)
                    };
                    _txtDiameter[r, c] = txt;
                    dr.Children.Add(txt);
                    dr.Children.Add(new TextBlock
                    {
                        Text = " mm",
                        FontSize = 9,
                        Foreground = new SolidColorBrush(COL_SUB),
                        VerticalAlignment = VerticalAlignment.Center
                    });
                    cs.Children.Add(dr);
                    cell.Child = cs;

                    Grid.SetRow(cell, r + 1);
                    Grid.SetColumn(cell, c + 1);
                    grid.Children.Add(cell);
                }
            }
            return grid;
        }

        // ═══════════════════════════════════════════════════════
        //  Bulk tools
        // ═══════════════════════════════════════════════════════

        private WrapPanel BuildBulkTools()
        {
            WrapPanel wrap = new WrapPanel();

            Button onBtn = ActionBtn("✔ Todos ON", COL_GREEN,
                Colors.White, 100);
            onBtn.Click += (s, e) => SetAll(true);
            wrap.Children.Add(onBtn);

            Button offBtn = ActionBtn("✖ Todos OFF", COL_RED,
                Colors.White, 100);
            offBtn.Click += (s, e) => SetAll(false);
            wrap.Children.Add(offBtn);

            wrap.Children.Add(new Border { Width = 16 });

            wrap.Children.Add(SmallLabel("Fila:"));
            ComboBox cmbR = new ComboBox
            {
                Width = 50,
                Height = 28,
                FontSize = 11,
                VerticalContentAlignment = VerticalAlignment.Center
            };
            for (int i = 1; i <= _rows; i++) cmbR.Items.Add(i);
            cmbR.SelectedIndex = 0;
            wrap.Children.Add(cmbR);
            Button trBtn = SmallBtn("Toggle");
            trBtn.Click += (s, e) =>
            {
                int ri = cmbR.SelectedIndex;
                for (int c = 0; c < _cols; c++)
                    _chkVisible[ri, c].IsChecked =
                        !(_chkVisible[ri, c].IsChecked == true);
            };
            wrap.Children.Add(trBtn);

            wrap.Children.Add(new Border { Width = 12 });

            wrap.Children.Add(SmallLabel("Col:"));
            ComboBox cmbC = new ComboBox
            {
                Width = 50,
                Height = 28,
                FontSize = 11,
                VerticalContentAlignment = VerticalAlignment.Center
            };
            for (int i = 1; i <= _cols; i++) cmbC.Items.Add(i);
            cmbC.SelectedIndex = 0;
            wrap.Children.Add(cmbC);
            Button tcBtn = SmallBtn("Toggle");
            tcBtn.Click += (s, e) =>
            {
                int ci = cmbC.SelectedIndex;
                for (int r = 0; r < _rows; r++)
                    _chkVisible[r, ci].IsChecked =
                        !(_chkVisible[r, ci].IsChecked == true);
            };
            wrap.Children.Add(tcBtn);

            wrap.Children.Add(new Border { Width = 16 });

            wrap.Children.Add(SmallLabel("ø para marcados:"));
            TextBox txBD = SmallInput(52);
            Border bdBD = MakeInputBorder();
            WireFocus(txBD, bdBD); bdBD.Child = txBD;
            wrap.Children.Add(bdBD);
            wrap.Children.Add(new TextBlock
            {
                Text = "mm",
                FontSize = 9,
                Foreground = new SolidColorBrush(COL_SUB),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0, 4, 0)
            });
            Button bdBtn = SmallBtn("Aplicar ø");
            bdBtn.Click += (s, e) =>
            {
                if (!double.TryParse(txBD.Text, out double d)) return;
                for (int r = 0; r < _rows; r++)
                    for (int c = 0; c < _cols; c++)
                        if (_chkVisible[r, c].IsChecked == true)
                            _txtDiameter[r, c].Text = d.ToString("F1");
            };
            wrap.Children.Add(bdBtn);

            return wrap;
        }

        // ═══════════════════════════════════════════════════════
        //  Action buttons
        // ═══════════════════════════════════════════════════════

        private StackPanel BuildActionButtons()
        {
            StackPanel row = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 4, 0, 0)
            };

            Button resetBtn = ActionBtn("Restaurar originales",
                COL_BTN_BG, COL_TEXT, 160);
            resetBtn.Click += (s, e) =>
            {
                ResultData = _original.Clone();
                ReloadUI();
            };
            row.Children.Add(resetBtn);

            Button cancelBtn = ActionBtn("Cancelar",
                COL_BTN_BG, COL_TEXT, 100);
            cancelBtn.Click += (s, e) =>
            {
                DialogResult = false; Close();
            };
            row.Children.Add(cancelBtn);

            Button applyBtn = ActionBtn(
                $"Aplicar a {_instanceCount} instancia(s)",
                COL_ACCENT, Colors.White, 240);
            applyBtn.FontSize = 14;
            applyBtn.FontWeight = FontWeights.SemiBold;
            applyBtn.Height = 40;
            applyBtn.Click += BtnApply_Click;
            row.Children.Add(applyBtn);

            return row;
        }

        // ═══════════════════════════════════════════════════════
        //  Apply
        // ═══════════════════════════════════════════════════════

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            bool err = false;

            if (double.TryParse(_txtAncho.Text, out double ancho))
                ResultData.Ancho = ancho;
            else err = MarkError(_txtAncho);

            if (double.TryParse(_txtAltura.Text, out double altura))
                ResultData.Altura = altura;
            else err = MarkError(_txtAltura);

            for (int i = 0; i < _cols; i++)
            {
                if (double.TryParse(_txtDistX[i].Text, out double dx))
                    ResultData.DistX[i] = dx;
                else err = MarkError(_txtDistX[i]);
            }
            for (int i = 0; i < _rows; i++)
            {
                if (double.TryParse(_txtDistY[i].Text, out double dy))
                    ResultData.DistY[i] = dy;
                else err = MarkError(_txtDistY[i]);
            }

            for (int r = 0; r < _rows; r++)
                for (int c = 0; c < _cols; c++)
                {
                    ResultData.Visible[r, c] =
                        _chkVisible[r, c].IsChecked == true;
                    if (double.TryParse(_txtDiameter[r, c].Text,
                        out double d))
                        ResultData.Diameter[r, c] = d;
                    else if (_chkVisible[r, c].IsChecked == true)
                        err = MarkError(_txtDiameter[r, c]);
                    else
                        ResultData.Diameter[r, c] = 0;
                }

            if (err)
            {
                MessageBox.Show(
                    "Campos inválidos (bordes rojos).",
                    "Error", MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
            Close();
        }

        private bool MarkError(TextBox tb)
        {
            tb.BorderBrush = new SolidColorBrush(COL_RED);
            tb.BorderThickness = new Thickness(2);
            return true;
        }

        // ═══════════════════════════════════════════════════════
        //  Helpers
        // ═══════════════════════════════════════════════════════

        private void SetAll(bool state)
        {
            for (int r = 0; r < _rows; r++)
                for (int c = 0; c < _cols; c++)
                    _chkVisible[r, c].IsChecked = state;
        }

        private void ReloadUI()
        {
            _txtAncho.Text = ResultData.Ancho.ToString("F1");
            _txtAltura.Text = ResultData.Altura.ToString("F1");
            for (int i = 0; i < _cols; i++)
                _txtDistX[i].Text = ResultData.DistX[i].ToString("F1");
            for (int i = 0; i < _rows; i++)
                _txtDistY[i].Text = ResultData.DistY[i].ToString("F1");
            for (int r = 0; r < _rows; r++)
                for (int c = 0; c < _cols; c++)
                {
                    _chkVisible[r, c].IsChecked =
                        ResultData.Visible[r, c];
                    _txtDiameter[r, c].Text =
                        ResultData.Diameter[r, c].ToString("F1");
                }
            DrawSchematic();
        }

        // ── UI Factories ──

        private Border MakeCard()
        {
            return new Border
            {
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(COL_BORDER),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Padding = new Thickness(10)
            };
        }

        private Border MakeInputBorder()
        {
            return new Border
            {
                CornerRadius = new CornerRadius(5),
                BorderBrush = new SolidColorBrush(COL_BORDER),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 6, 0)
            };
        }

        private void WireFocus(TextBox txt, Border bdr)
        {
            txt.GotFocus += (s, e) =>
                bdr.BorderBrush = new SolidColorBrush(COL_ACCENT);
            txt.LostFocus += (s, e) =>
                bdr.BorderBrush = new SolidColorBrush(COL_BORDER);
        }

        private TextBlock SectionHeader(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(COL_ACCENT),
                Margin = new Thickness(0, 8, 0, 2)
            };
        }

        private TextBlock SubLabel(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = 11,
                Foreground = new SolidColorBrush(COL_SUB),
                Margin = new Thickness(0, 0, 0, 2)
            };
        }

        private TextBlock SmallLabel(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = 11,
                Foreground = new SolidColorBrush(COL_TEXT),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 6, 0)
            };
        }

        private TextBox SmallInput(double width)
        {
            return new TextBox
            {
                Width = width,
                Height = 26,
                FontSize = 11,
                TextAlignment = TextAlignment.Center,
                VerticalContentAlignment = VerticalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(2, 0, 2, 0)
            };
        }

        private Button SmallBtn(string text)
        {
            Button btn = new Button
            {
                Content = text,
                Height = 28,
                FontSize = 11,
                Padding = new Thickness(10, 0, 10, 0),
                Margin = new Thickness(0, 0, 6, 0),
                Cursor = Cursors.Hand,
                Background = new SolidColorBrush(COL_BTN_BG),
                Foreground = new SolidColorBrush(COL_TEXT),
                BorderThickness = new Thickness(0)
            };
            btn.Template = RoundBtnTemplate(COL_BTN_BG);
            return btn;
        }

        private Button ActionBtn(string text, Color bg,
            Color fg, double width)
        {
            Button btn = new Button
            {
                Content = text,
                Width = width,
                Height = 36,
                FontSize = 13,
                Margin = new Thickness(0, 0, 8, 0),
                Cursor = Cursors.Hand,
                Foreground = new SolidColorBrush(fg),
                Background = new SolidColorBrush(bg),
                BorderThickness = new Thickness(0)
            };
            btn.Template = RoundBtnTemplate(bg);
            return btn;
        }

        private ControlTemplate RoundBtnTemplate(Color bg)
        {
            ControlTemplate t = new ControlTemplate(typeof(Button));
            FrameworkElementFactory bdr =
                new FrameworkElementFactory(typeof(Border));
            bdr.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
            bdr.SetValue(Border.BackgroundProperty, new SolidColorBrush(bg));
            bdr.SetValue(Border.PaddingProperty, new Thickness(14, 6, 14, 6));
            FrameworkElementFactory cp =
                new FrameworkElementFactory(typeof(ContentPresenter));
            cp.SetValue(ContentPresenter.HorizontalAlignmentProperty,
                HorizontalAlignment.Center);
            cp.SetValue(ContentPresenter.VerticalAlignmentProperty,
                VerticalAlignment.Center);
            bdr.AppendChild(cp);
            t.VisualTree = bdr;
            return t;
        }
    }
}