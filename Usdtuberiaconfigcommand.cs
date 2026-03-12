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

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            try
            {
                ICollection<ElementId> selIds = uidoc.Selection.GetElementIds();
                List<FamilyInstance> instances = new List<FamilyInstance>();
                foreach (ElementId id in selIds)
                { Element el = doc.GetElement(id); if (el is FamilyInstance fi) instances.Add(fi); }
                if (instances.Count == 0)
                { TaskDialog.Show("HMV - Ductos Editor", "Seleccione al menos una instancia USD."); return Result.Cancelled; }

                FamilyInstance primary = instances[0];
                USDMatrixSizeWindow sizeWin = new USDMatrixSizeWindow();
                if (sizeWin.ShowDialog() != true) return Result.Cancelled;
                int rows = sizeWin.SelectedRows, cols = sizeWin.SelectedCols;

                string diaPrefix = FindDiameterPrefix(primary);
                USDTuberiaData data = ReadParameters(primary, diaPrefix);
                if (data == null) return Result.Failed;

                string typeInfo = $"Tipo actual: {primary.Symbol.Name}";
                USDTuberiaWindow win = new USDTuberiaWindow(data, instances.Count, typeInfo, rows, cols);
                if (win.ShowDialog() != true) return Result.Cancelled;

                USDTuberiaData newData = win.ResultData;
                for (int r = 0; r < MAX_ROWS; r++)
                    for (int c = 0; c < MAX_COLS; c++)
                        if (r >= rows || c >= cols) newData.Visible[r, c] = false;

                string targetName = BuildTypeName(newData, rows, cols);
                using (Transaction tx = new Transaction(doc, "HMV - Configurar USD Tuberías"))
                {
                    tx.Start();
                    FamilySymbol targetType = FindOrCreateType(doc, primary, targetName);
                    if (!targetType.IsActive) targetType.Activate();
                    WriteParametersToType(targetType, newData, diaPrefix);
                    foreach (FamilyInstance fi in instances)
                        if (fi.Symbol.Id != targetType.Id) fi.ChangeTypeId(targetType.Id);
                    tx.Commit();
                }
                PurgeUnusedTypes(doc, primary.Symbol.Family);
                TaskDialog.Show("HMV - Ductos Editor",
                    $"Configuración aplicada.\nTipo: {targetName}\nInstancias: {instances.Count}  |  Matriz: {rows}×{cols}");
                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }
            catch (Exception ex) { message = ex.Message; return Result.Failed; }
        }

        #region Type naming
        string BuildTypeName(USDTuberiaData d, int rows, int cols)
        {
            int n = 0; List<double> dias = new List<double>();
            for (int r = 0; r < rows; r++) for (int c = 0; c < cols; c++)
                    if (d.Visible[r, c]) { n++; double v = d.Diameter[r, c]; if (v > 0 && !dias.Any(x => Math.Abs(x - v) < 0.1)) dias.Add(v); }
            dias.Sort();
            string ds = dias.Count == 0 ? "0\"" : dias.Count == 1 ? Inch(dias[0]) : "(" + string.Join(",", dias.Select(Inch)) + ")";
            return $"{n} \u00f8 {ds} ({d.Ancho / 1000.0:F2} m x {d.Altura / 1000.0:F2} m)";
        }
        string Inch(double mm) { double v = mm / 25.4; return Math.Abs(v - Math.Round(v)) < 0.05 ? $"{Math.Round(v):F0}\"" : $"{v:F1}\""; }
        #endregion

        #region Type management
        FamilySymbol FindOrCreateType(Document doc, FamilyInstance src, string name)
        {
            Family fam = src.Symbol.Family;
            foreach (ElementId id in fam.GetFamilySymbolIds())
            { FamilySymbol fs = doc.GetElement(id) as FamilySymbol; if (fs != null && fs.Name == name) return fs; }
            return doc.GetElement(src.Symbol.Duplicate(name).Id) as FamilySymbol;
        }
        void PurgeUnusedTypes(Document doc, Family fam)
        {
            var ids = fam.GetFamilySymbolIds(); if (ids.Count <= 1) return;
            List<FamilySymbol> unused = new List<FamilySymbol>();
            foreach (ElementId id in ids)
            { FamilySymbol fs = doc.GetElement(id) as FamilySymbol; if (fs == null) continue; int cnt = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Count(fi => fi.Symbol.Id == id); if (cnt == 0) unused.Add(fs); }
            if (unused.Count == 0) return;
            string names = string.Join("\n", unused.Select(u => "  • " + u.Name));
            if (MessageBox.Show($"{unused.Count} tipo(s) sin instancias:\n\n{names}\n\n¿Eliminarlos?", "HMV - Purgar", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
            using (Transaction tx = new Transaction(doc, "HMV - Purgar tipos"))
            { tx.Start(); foreach (var fs in unused) try { doc.Delete(fs.Id); } catch { } tx.Commit(); }
        }
        #endregion

        #region Parameter IO
        string FindDiameterPrefix(FamilyInstance fi)
        {
            FamilySymbol sym = fi.Symbol;
            string[] cands = { "HMV_USD_\u00f8 Tuberia ", "HMV_USD_\u00d8 Tuberia ", "HMV_USD_o Tuberia ", "HMV_USD_O Tuberia ", "HMV_USD_\u00f8 TUBERIA ", "HMV_USD_\u00d8 TUBERIA ", "HMV_USD_\u2300 Tuberia " };
            foreach (string p in cands) if (sym.LookupParameter(p + "1-1") != null) return p;
            foreach (Parameter p in sym.Parameters) { string n = p.Definition?.Name ?? ""; if (n.Contains("Tuberia 1-1") && !n.Contains("Visibilidad") && !n.Contains("Distancia")) { int i = n.IndexOf("1-1"); if (i > 0) return n.Substring(0, i); } }
            return "HMV_USD_\u00f8 Tuberia ";
        }
        USDTuberiaData ReadParameters(FamilyInstance fi, string dp)
        {
            USDTuberiaData d = new USDTuberiaData(); List<string> miss = new List<string>();
            d.Ancho = TryGD(fi, "HMV_USD_Ancho", ref miss); d.Altura = TryGD(fi, "HMV_USD_Altura", ref miss);
            for (int i = 1; i <= MAX_DIST; i++) { d.DistX[i - 1] = TryGD(fi, $"HMV_USD_Distancia {i}_X", ref miss); d.DistY[i - 1] = TryGD(fi, $"HMV_USD_Distancia {i}_Y", ref miss); }
            for (int r = 1; r <= MAX_ROWS; r++) for (int c = 1; c <= MAX_COLS; c++) { d.Visible[r - 1, c - 1] = TryGB(fi, $"HMV_USD_Visibilidad Tuberia {r}-{c}", ref miss); d.Diameter[r - 1, c - 1] = TryGD(fi, $"{dp}{r}-{c}", ref miss); }
            if (miss.Count > 20) { TaskDialog td = new TaskDialog("HMV - Diagnóstico"); td.MainInstruction = $"{miss.Count} parámetros faltantes."; td.MainContent = string.Join("\n", miss.Take(8).Select(m => "  • " + m)); td.Show(); return null; }
            if (miss.Count > 0) { TaskDialog td = new TaskDialog("HMV - Aviso"); td.MainInstruction = $"{miss.Count} parámetro(s) no encontrado(s)."; td.MainContent = string.Join("\n", miss.Take(10).Select(m => "  • " + m)); td.Show(); }
            return d;
        }
        void WriteParametersToType(FamilySymbol sym, USDTuberiaData d, string dp)
        {
            SD(sym, "HMV_USD_Ancho", d.Ancho); SD(sym, "HMV_USD_Altura", d.Altura);
            for (int i = 1; i <= MAX_DIST; i++) { SD(sym, $"HMV_USD_Distancia {i}_X", d.DistX[i - 1]); SD(sym, $"HMV_USD_Distancia {i}_Y", d.DistY[i - 1]); }
            for (int r = 1; r <= MAX_ROWS; r++) for (int c = 1; c <= MAX_COLS; c++) { SB(sym, $"HMV_USD_Visibilidad Tuberia {r}-{c}", d.Visible[r - 1, c - 1]); SD(sym, $"{dp}{r}-{c}", d.Diameter[r - 1, c - 1]); }
        }
        double TryGD(FamilyInstance fi, string n, ref List<string> m) { var p = fi.Symbol.LookupParameter(n); if (p == null) { m.Add(n); return 0; } return p.AsDouble() * 304.8; }
        bool TryGB(FamilyInstance fi, string n, ref List<string> m) { var p = fi.Symbol.LookupParameter(n); if (p == null) { m.Add(n); return false; } return p.AsInteger() == 1; }
        void SD(FamilySymbol s, string n, double mm) { var p = s.LookupParameter(n); if (p != null && !p.IsReadOnly) p.Set(mm / 304.8); }
        void SB(FamilySymbol s, string n, bool v) { var p = s.LookupParameter(n); if (p != null && !p.IsReadOnly) p.Set(v ? 1 : 0); }
        #endregion
    }

    public class USDTuberiaData
    {
        public double Ancho, Altura;
        public double[] DistX = new double[USDTuberiaConfigCommand.MAX_DIST];
        public double[] DistY = new double[USDTuberiaConfigCommand.MAX_DIST];
        public bool[,] Visible = new bool[USDTuberiaConfigCommand.MAX_ROWS, USDTuberiaConfigCommand.MAX_COLS];
        public double[,] Diameter = new double[USDTuberiaConfigCommand.MAX_ROWS, USDTuberiaConfigCommand.MAX_COLS];
        public USDTuberiaData Clone()
        {
            USDTuberiaData c = new USDTuberiaData { Ancho = Ancho, Altura = Altura };
            Array.Copy(DistX, c.DistX, DistX.Length); Array.Copy(DistY, c.DistY, DistY.Length);
            Array.Copy(Visible, c.Visible, Visible.Length); Array.Copy(Diameter, c.Diameter, Diameter.Length);
            return c;
        }
    }

    // ═══════════════════════════════════════════════════════════
    //  Matrix size selector
    // ═══════════════════════════════════════════════════════════
    public class USDMatrixSizeWindow : Window
    {
        public int SelectedRows { get; private set; } = 5;
        public int SelectedCols { get; private set; } = 9;
        private ComboBox _cmbR, _cmbC; private Canvas _prev;
        static readonly Color CA = Color.FromRgb(0, 120, 212), CG = Color.FromRgb(40, 167, 69), CT = Color.FromRgb(30, 30, 30), CS = Color.FromRgb(120, 120, 130), CB = Color.FromRgb(200, 200, 210);

        public USDMatrixSizeWindow()
        {
            Title = "HMV Tools - Ductos Editor"; Width = 480; Height = 520;
            WindowStartupLocation = WindowStartupLocation.CenterScreen; ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(Color.FromRgb(245, 245, 248));
            StackPanel root = new StackPanel { Margin = new Thickness(24) }; Content = root;
            root.Children.Add(new TextBlock { Text = "¿De cuánto por cuánto es la\nmatriz de ductos?", FontSize = 20, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(CT), Margin = new Thickness(0, 0, 0, 6), TextWrapping = TextWrapping.Wrap });
            root.Children.Add(new TextBlock { Text = "Seleccione filas y columnas.", FontSize = 12, Foreground = new SolidColorBrush(CS), Margin = new Thickness(0, 0, 0, 20) });
            Grid sg = new Grid { Margin = new Thickness(0, 0, 0, 20) };
            sg.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            sg.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) });
            sg.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            root.Children.Add(sg);
            StackPanel rp = new StackPanel();
            rp.Children.Add(new TextBlock { Text = "Filas (Y)", FontSize = 12, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(CA), Margin = new Thickness(0, 0, 0, 6) });
            _cmbR = new ComboBox { Height = 36, FontSize = 14, VerticalContentAlignment = VerticalAlignment.Center };
            for (int i = 1; i <= 5; i++) _cmbR.Items.Add(i); _cmbR.SelectedIndex = 4; _cmbR.SelectionChanged += (s, e) => DrawP(); rp.Children.Add(_cmbR);
            Grid.SetColumn(rp, 0); sg.Children.Add(rp);
            var xl = new TextBlock { Text = "×", FontSize = 22, FontWeight = FontWeights.Bold, Foreground = new SolidColorBrush(CT), HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Bottom, Margin = new Thickness(0, 0, 0, 6) };
            Grid.SetColumn(xl, 1); sg.Children.Add(xl);
            StackPanel cp = new StackPanel();
            cp.Children.Add(new TextBlock { Text = "Columnas (X)", FontSize = 12, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(CA), Margin = new Thickness(0, 0, 0, 6) });
            _cmbC = new ComboBox { Height = 36, FontSize = 14, VerticalContentAlignment = VerticalAlignment.Center };
            for (int i = 1; i <= 9; i++) _cmbC.Items.Add(i); _cmbC.SelectedIndex = 8; _cmbC.SelectionChanged += (s, e) => DrawP(); cp.Children.Add(_cmbC);
            Grid.SetColumn(cp, 2); sg.Children.Add(cp);
            Border pb = new Border { CornerRadius = new CornerRadius(8), BorderBrush = new SolidColorBrush(CB), BorderThickness = new Thickness(1), Background = Brushes.White, Height = 160, Margin = new Thickness(0, 0, 0, 20) };
            _prev = new Canvas { ClipToBounds = true }; _prev.Loaded += (s, e) => DrawP(); pb.Child = _prev; root.Children.Add(pb);
            StackPanel br = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            Button cb2 = MkB("Cancelar", Color.FromRgb(240, 240, 243), CT, 100); cb2.Click += (s, e) => { DialogResult = false; Close(); }; br.Children.Add(cb2);
            Button ob = MkB("Continuar →", CA, Colors.White, 140); ob.FontWeight = FontWeights.SemiBold; ob.Click += (s, e) => { SelectedRows = (int)_cmbR.SelectedItem; SelectedCols = (int)_cmbC.SelectedItem; DialogResult = true; Close(); }; br.Children.Add(ob);
            root.Children.Add(br);
        }
        void DrawP()
        {
            _prev.Children.Clear(); double w = _prev.ActualWidth, h = _prev.ActualHeight; if (w < 10) return;
            int r = (int)(_cmbR.SelectedItem ?? 5), c = (int)(_cmbC.SelectedItem ?? 9);
            double cW = (w - 50) / Math.Max(c, 1), cH = (h - 40) / Math.Max(r, 1), rad = Math.Min(Math.Min(cW, cH) * 0.35, 12);
            var lbl = new TextBlock { Text = $"Matriz {r} × {c} = {r * c} posiciones", FontSize = 10, Foreground = new SolidColorBrush(CS) };
            Canvas.SetLeft(lbl, 4); Canvas.SetTop(lbl, 2); _prev.Children.Add(lbl);
            for (int ri = 0; ri < r; ri++) for (int ci = 0; ci < c; ci++)
                { double cx = 30 + ci * cW + cW / 2, cy = 20 + ri * cH + cH / 2; var el = new Ellipse { Width = rad * 2, Height = rad * 2, Stroke = new SolidColorBrush(CG), StrokeThickness = 1.5, Fill = new SolidColorBrush(Color.FromArgb(80, 40, 167, 69)) }; Canvas.SetLeft(el, cx - rad); Canvas.SetTop(el, cy - rad); _prev.Children.Add(el); }
        }
        Button MkB(string t, Color bg, Color fg, double w)
        {
            Button b = new Button { Content = t, Width = w, Height = 36, FontSize = 13, Margin = new Thickness(0, 0, 8, 0), Cursor = Cursors.Hand, Foreground = new SolidColorBrush(fg), Background = new SolidColorBrush(bg), BorderThickness = new Thickness(0) };
            var tp = new ControlTemplate(typeof(Button)); var bd = new FrameworkElementFactory(typeof(Border));
            bd.SetValue(Border.CornerRadiusProperty, new CornerRadius(6)); bd.SetValue(Border.BackgroundProperty, new SolidColorBrush(bg)); bd.SetValue(Border.PaddingProperty, new Thickness(14, 6, 14, 6));
            var cp2 = new FrameworkElementFactory(typeof(ContentPresenter)); cp2.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center); cp2.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            bd.AppendChild(cp2); tp.VisualTree = bd; b.Template = tp; return b;
        }
    }

    // ═══════════════════════════════════════════════════════════
    //  Main editor — FULL INTERACTIVE SCHEMATIC
    // ═══════════════════════════════════════════════════════════
    public class USDTuberiaWindow : Window
    {
        public USDTuberiaData ResultData { get; private set; }
        private readonly USDTuberiaData _original;
        private readonly int _instanceCount, _rows, _cols;
        private readonly string _typeInfo;
        private int _dxCount, _dyCount;

        private TextBox _txtAncho, _txtAltura;
        private TextBox[] _txtDistX, _txtDistY;
        private TextBox _txtTodosX, _txtTodosY;
        private CheckBox[,] _chkVisible;
        private TextBox[,] _txtDiameter;
        private Ellipse[,] _circles;

        static readonly Color COL_BG = Color.FromRgb(245, 245, 248);
        static readonly Color COL_ACCENT = Color.FromRgb(0, 120, 212);
        static readonly Color COL_GREEN = Color.FromRgb(40, 167, 69);
        static readonly Color COL_RED = Color.FromRgb(220, 53, 69);
        static readonly Color COL_BORDER = Color.FromRgb(200, 200, 210);
        static readonly Color COL_TEXT = Color.FromRgb(30, 30, 30);
        static readonly Color COL_SUB = Color.FromRgb(120, 120, 130);
        static readonly Color COL_BTN = Color.FromRgb(240, 240, 243);

        static readonly SolidColorBrush BR_ON = new SolidColorBrush(Color.FromArgb(180, 40, 167, 69));
        static readonly SolidColorBrush BR_PEND = new SolidColorBrush(Color.FromArgb(200, 255, 193, 7));
        static readonly SolidColorBrush BR_OFF = new SolidColorBrush(Color.FromArgb(40, 180, 180, 180));
        static readonly SolidColorBrush BR_STR_ON = new SolidColorBrush(COL_GREEN);
        static readonly SolidColorBrush BR_STR_PEND = new SolidColorBrush(Color.FromRgb(230, 170, 0));
        static readonly SolidColorBrush BR_STR_OFF = new SolidColorBrush(Color.FromRgb(190, 190, 190));

        // Grid layout helpers
        // Columns: [0:Ylabels] [1:BordeX gap] [2:Hole0] [3:gap] [4:Hole1] ... [2c+2:HoleC] [last-1:lastGap?] [last:TodosX]
        // Formula: GapCol(d) = 1 + d*2 for d=0.._dxCount-1, HoleCol(c) = 2 + c*2
        int GapCol(int d) => 1 + d * 2;
        int HoleCol(int c) => 2 + c * 2;
        int GapRow(int d) => 1 + d * 2;
        int HoleRow(int r) => 2 + r * 2;
        int _totalCols, _totalRows;

        public USDTuberiaWindow(USDTuberiaData data, int instCount, string typeInfo, int rows, int cols)
        {
            _original = data; _instanceCount = instCount; _typeInfo = typeInfo;
            _rows = rows; _cols = cols;
            ResultData = data.Clone();
            _dxCount = Math.Min(_cols + 1, USDTuberiaConfigCommand.MAX_DIST);
            _dyCount = Math.Min(_rows + 1, USDTuberiaConfigCommand.MAX_DIST);
            _txtDistX = new TextBox[_dxCount]; _txtDistY = new TextBox[_dyCount];
            _chkVisible = new CheckBox[_rows, _cols];
            _txtDiameter = new TextBox[_rows, _cols];
            _circles = new Ellipse[_rows, _cols];

            // Total grid size: labels + gaps + holes + TodosCol
            // Gaps: _dxCount (including bordeX and possibly one after last hole)
            // Holes: _cols
            // Total cols = 1 + _dxCount + _cols + 1
            // But gaps interleave with holes: col0=labels, then alternating gap/hole
            // If _dxCount > _cols: there's a trailing gap after last hole
            _totalCols = 1 + _cols + _dxCount + 1;  // labels + holes + gaps + TodosX
            _totalRows = 1 + _rows + _dyCount + 1;

            BuildUI();
        }

        void BuildUI()
        {
            Title = "HMV Tools - Ductos Editor";
            Width = Math.Max(800, 100 + (_cols + _dxCount) * 65);
            Height = Math.Max(620, 250 + (_rows + _dyCount) * 55);
            MinWidth = 750; MinHeight = 580;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Background = new SolidColorBrush(COL_BG);
            ResizeMode = ResizeMode.CanResize;

            ScrollViewer scr = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                Padding = new Thickness(16)
            };
            Content = scr;
            StackPanel root = new StackPanel();
            scr.Content = root;

            // ── Header ──
            root.Children.Add(new TextBlock { Text = $"Ductos Editor  —  Matriz {_rows} × {_cols}", FontSize = 18, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(COL_TEXT), Margin = new Thickness(0, 0, 0, 2) });
            root.Children.Add(new TextBlock { Text = $"{_instanceCount} instancia(s)  ·  {_typeInfo}", FontSize = 11, Foreground = new SolidColorBrush(COL_SUB), Margin = new Thickness(0, 0, 0, 6), TextWrapping = TextWrapping.Wrap });

            // ── Legend ──
            WrapPanel legend = new WrapPanel { Margin = new Thickness(0, 0, 0, 8) };
            legend.Children.Add(LegendDot(BR_ON, "Activo (original)"));
            legend.Children.Add(LegendDot(BR_PEND, "Cambio pendiente"));
            legend.Children.Add(LegendDot(BR_OFF, "Inactivo"));
            root.Children.Add(legend);

            // ══════════════════════════════════════════════
            //  INTERACTIVE SCHEMATIC WITH DISTANCES
            // ══════════════════════════════════════════════
            Border schemCard = MkCard();
            schemCard.Padding = new Thickness(8);
            schemCard.Margin = new Thickness(0, 0, 0, 10);

            Grid g = new Grid();

            // Define columns
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(60) }); // col 0: Y labels
            for (int d = 0; d < _dxCount; d++)
            {
                g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(55) }); // gap col
                if (d < _cols)
                    g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star), MinWidth = 88 }); // hole col
            }
            // If _dxCount > _cols, last gap has no hole after it — already added
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) }); // TodosX col

            // Define rows
            g.RowDefinitions.Add(new RowDefinition { Height = new GridLength(60) }); // row 0: X labels
            for (int d = 0; d < _dyCount; d++)
            {
                g.RowDefinitions.Add(new RowDefinition { Height = new GridLength(30) }); // gap row
                if (d < _rows)
                    g.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto, MinHeight = 88 }); // hole row
            }
            g.RowDefinitions.Add(new RowDefinition { Height = new GridLength(60) }); // TodosY row

            // ── RED BASELINE: 0,0 corner ──
            TextBlock zeroLabel = new TextBlock
            {
                Text = "0,0",
                FontSize = 10,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(COL_RED),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Bottom,
                Margin = new Thickness(0, 0, 2, 0)
            };
            Grid.SetRow(zeroLabel, 0); Grid.SetColumn(zeroLabel, 0);
            g.Children.Add(zeroLabel);

            // Red solid horizontal line
            Line redTop = new Line
            {
                Stroke = new SolidColorBrush(COL_RED),
                StrokeThickness = 2,
                X1 = 0, Y1 = 0, X2 = 1, Y2 = 0,
                Stretch = Stretch.Fill,
                VerticalAlignment = VerticalAlignment.Bottom,
                HorizontalAlignment = HorizontalAlignment.Stretch
            };
            Grid.SetRow(redTop, 0);
            Grid.SetColumn(redTop, 0);
            Grid.SetColumnSpan(redTop, g.ColumnDefinitions.Count - 1);
            g.Children.Add(redTop);

            // Red solid vertical line
            Line redLeft = new Line
            {
                Stroke = new SolidColorBrush(COL_RED),
                StrokeThickness = 2,
                X1 = 0, Y1 = 0, X2 = 0, Y2 = 1,
                Stretch = Stretch.Fill,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Stretch
            };
            Grid.SetRow(redLeft, 0);
            Grid.SetColumn(redLeft, 0);
            Grid.SetRowSpan(redLeft, g.RowDefinitions.Count - 1);
            g.Children.Add(redLeft);

            // ── X DISTANCE INPUTS (in gap columns, row 0) ──
            for (int d = 0; d < _dxCount; d++)
            {
                int gc = GapCol(d);
                string label = (d == 0 || d == _dxCount - 1) ? "Borde X" : $"dX{d + 1}";
                StackPanel sp = new StackPanel { VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center };
                sp.Children.Add(new TextBlock { Text = label, FontSize = 7.5, Foreground = new SolidColorBrush(COL_ACCENT), HorizontalAlignment = HorizontalAlignment.Center, FontWeight = FontWeights.SemiBold });
                TextBox tb = new TextBox { Width = 44, Height = 18, FontSize = 8.5, Text = ResultData.DistX[d].ToString("F1"), TextAlignment = TextAlignment.Center, VerticalContentAlignment = VerticalAlignment.Center, BorderThickness = new Thickness(1), BorderBrush = new SolidColorBrush(COL_BORDER), Padding = new Thickness(1, 0, 1, 0) };
                _txtDistX[d] = tb;
                sp.Children.Add(tb);
                Grid.SetRow(sp, 0); Grid.SetColumn(sp, gc);
                g.Children.Add(sp);
            }

            // ── X HOLE POSITION LABELS (in hole columns, row 0) ──
            for (int c = 0; c < _cols; c++)
            {
                TextBlock hdr = new TextBlock { Text = $"X{c + 1}", FontSize = 10, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(COL_ACCENT), HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Bottom };
                Grid.SetRow(hdr, 0); Grid.SetColumn(hdr, HoleCol(c));
                g.Children.Add(hdr);
            }

            // ── Y DISTANCE INPUTS (in gap rows, col 0) ──
            for (int d = 0; d < _dyCount; d++)
            {
                int gr = GapRow(d);
                string label = (d == 0 || d == _dyCount - 1) ? "Borde Y" : $"dY{d + 1}";
                StackPanel sp = new StackPanel { VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center };
                sp.Children.Add(new TextBlock { Text = label, FontSize = 7.5, Foreground = new SolidColorBrush(COL_ACCENT), HorizontalAlignment = HorizontalAlignment.Center, FontWeight = FontWeights.SemiBold });
                TextBox tb = new TextBox { Width = 44, Height = 18, FontSize = 8.5, Text = ResultData.DistY[d].ToString("F1"), TextAlignment = TextAlignment.Center, VerticalContentAlignment = VerticalAlignment.Center, BorderThickness = new Thickness(1), BorderBrush = new SolidColorBrush(COL_BORDER), Padding = new Thickness(1, 0, 1, 0) };
                _txtDistY[d] = tb;
                sp.Children.Add(tb);
                Grid.SetRow(sp, gr); Grid.SetColumn(sp, 0);
                g.Children.Add(sp);
            }

            // ── Y HOLE POSITION LABELS (in hole rows, col 0) ──
            for (int r = 0; r < _rows; r++)
            {
                TextBlock lbl = new TextBlock { Text = $"Y{r + 1}", FontSize = 10, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(COL_ACCENT), HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
                Grid.SetRow(lbl, HoleRow(r)); Grid.SetColumn(lbl, 0);
                g.Children.Add(lbl);
            }

            // ── TODOS X (top-right corner) ──
            int todosXCol = g.ColumnDefinitions.Count - 1;
            StackPanel txSp = new StackPanel { VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center };
            txSp.Children.Add(new TextBlock { Text = "Todos X", FontSize = 7.5, Foreground = new SolidColorBrush(COL_ACCENT), HorizontalAlignment = HorizontalAlignment.Center, FontWeight = FontWeights.SemiBold });
            _txtTodosX = new TextBox { Width = 44, Height = 18, FontSize = 8.5, TextAlignment = TextAlignment.Center, VerticalContentAlignment = VerticalAlignment.Center, BorderThickness = new Thickness(1), BorderBrush = new SolidColorBrush(COL_ACCENT), Padding = new Thickness(1, 0, 1, 0) };
            txSp.Children.Add(_txtTodosX);
            Button btnTX = new Button { Content = "▶", FontSize = 8, Width = 22, Height = 16, Margin = new Thickness(0, 2, 0, 0), Cursor = Cursors.Hand, Background = new SolidColorBrush(COL_ACCENT), Foreground = Brushes.White, BorderThickness = new Thickness(0) };
            btnTX.Click += (s, e) => { if (double.TryParse(_txtTodosX.Text, out double v)) foreach (var t in _txtDistX) t.Text = v.ToString("F1"); };
            txSp.Children.Add(btnTX);
            Grid.SetRow(txSp, 0); Grid.SetColumn(txSp, todosXCol);
            g.Children.Add(txSp);

            // ── TODOS Y (bottom-left corner) ──
            int todosYRow = g.RowDefinitions.Count - 1;
            StackPanel tySp = new StackPanel { VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center };
            tySp.Children.Add(new TextBlock { Text = "Todos Y", FontSize = 7.5, Foreground = new SolidColorBrush(COL_ACCENT), HorizontalAlignment = HorizontalAlignment.Center, FontWeight = FontWeights.SemiBold });
            _txtTodosY = new TextBox { Width = 44, Height = 18, FontSize = 8.5, TextAlignment = TextAlignment.Center, VerticalContentAlignment = VerticalAlignment.Center, BorderThickness = new Thickness(1), BorderBrush = new SolidColorBrush(COL_ACCENT), Padding = new Thickness(1, 0, 1, 0) };
            tySp.Children.Add(_txtTodosY);
            Button btnTY = new Button { Content = "▶", FontSize = 8, Width = 22, Height = 16, Margin = new Thickness(0, 2, 0, 0), Cursor = Cursors.Hand, Background = new SolidColorBrush(COL_ACCENT), Foreground = Brushes.White, BorderThickness = new Thickness(0) };
            btnTY.Click += (s, e) => { if (double.TryParse(_txtTodosY.Text, out double v)) foreach (var t in _txtDistY) t.Text = v.ToString("F1"); };
            tySp.Children.Add(btnTY);
            Grid.SetRow(tySp, todosYRow); Grid.SetColumn(tySp, 0);
            g.Children.Add(tySp);

            // ── HOLE CIRCLES (in hole cells) ──
            for (int r = 0; r < _rows; r++)
                for (int c = 0; c < _cols; c++)
                {
                    int row = r, col = c;
                    bool vis = ResultData.Visible[r, c];

                    Grid cell = new Grid { Margin = new Thickness(2), MinWidth = 82, MinHeight = 82 };

                    Ellipse circle = new Ellipse { Width = 82, Height = 82, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, Fill = vis ? BR_ON : BR_OFF, Stroke = vis ? BR_STR_ON : BR_STR_OFF, StrokeThickness = 2.5 };
                    _circles[r, c] = circle;
                    cell.Children.Add(circle);

                    StackPanel content = new StackPanel { VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center };
                    CheckBox chk = new CheckBox { Content = $"{r + 1}-{c + 1}", FontSize = 10, FontWeight = FontWeights.SemiBold, IsChecked = vis, Foreground = new SolidColorBrush(vis ? Colors.White : Color.FromRgb(100, 100, 100)), HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 0, 0, 2) };
                    chk.Checked += (s, e) => UpdHole(row, col, true);
                    chk.Unchecked += (s, e) => UpdHole(row, col, false);
                    _chkVisible[r, c] = chk;
                    content.Children.Add(chk);

                    TextBox dia = new TextBox { Width = 48, Height = 18, FontSize = 9.5, Text = ResultData.Diameter[r, c].ToString("F1"), TextAlignment = TextAlignment.Center, VerticalContentAlignment = VerticalAlignment.Center, BorderThickness = new Thickness(1), BorderBrush = new SolidColorBrush(Color.FromArgb(120, 255, 255, 255)), Background = new SolidColorBrush(Color.FromArgb(vis ? (byte)180 : (byte)60, 255, 255, 255)), IsEnabled = vis, Padding = new Thickness(2, 0, 2, 0), HorizontalAlignment = HorizontalAlignment.Center };
                    dia.LostFocus += DiaLostFocus;
                    _txtDiameter[r, c] = dia;
                    content.Children.Add(dia);
                    content.Children.Add(new TextBlock { Text = "mm/\"", FontSize = 7.5, Foreground = new SolidColorBrush(vis ? Colors.White : Color.FromRgb(160, 160, 160)), HorizontalAlignment = HorizontalAlignment.Center });

                    cell.Children.Add(content);
                    Grid.SetRow(cell, HoleRow(r));
                    Grid.SetColumn(cell, HoleCol(c));
                    g.Children.Add(cell);
                }

            // ── ROW DIAMETER CONTROLS (TodosX column, one per row) ──
            for (int r = 0; r < _rows; r++)
            {
                int rowIdx = r;
                StackPanel rsp = new StackPanel { VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center };
                rsp.Children.Add(new TextBlock { Text = $"Fila {r + 1}", FontSize = 7.5, Foreground = new SolidColorBrush(COL_ACCENT), HorizontalAlignment = HorizontalAlignment.Center, FontWeight = FontWeights.SemiBold });
                TextBox rtb = new TextBox { Width = 48, Height = 18, FontSize = 8.5, TextAlignment = TextAlignment.Center, VerticalContentAlignment = VerticalAlignment.Center, BorderThickness = new Thickness(1), BorderBrush = new SolidColorBrush(COL_ACCENT), Padding = new Thickness(1, 0, 1, 0) };
                rtb.LostFocus += DiaLostFocus;
                rsp.Children.Add(rtb);
                rsp.Children.Add(new TextBlock { Text = "mm/\"", FontSize = 7, Foreground = new SolidColorBrush(COL_SUB), HorizontalAlignment = HorizontalAlignment.Center });
                Button rbtn = new Button { Content = "▶", FontSize = 8, Width = 22, Height = 16, Margin = new Thickness(0, 2, 0, 0), Cursor = Cursors.Hand, Background = new SolidColorBrush(COL_ACCENT), Foreground = Brushes.White, BorderThickness = new Thickness(0) };
                rbtn.Click += (s, e) => { if (ParseDia(rtb.Text, out double v)) for (int ci = 0; ci < _cols; ci++) _txtDiameter[rowIdx, ci].Text = v.ToString("F1"); };
                rsp.Children.Add(rbtn);
                Grid.SetRow(rsp, HoleRow(r)); Grid.SetColumn(rsp, todosXCol);
                g.Children.Add(rsp);
            }

            // ── COLUMN DIAMETER CONTROLS (TodosY row, one per column) ──
            for (int c = 0; c < _cols; c++)
            {
                int colIdx = c;
                StackPanel csp = new StackPanel { VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center };
                csp.Children.Add(new TextBlock { Text = $"Col {c + 1}", FontSize = 7.5, Foreground = new SolidColorBrush(COL_ACCENT), HorizontalAlignment = HorizontalAlignment.Center, FontWeight = FontWeights.SemiBold });
                TextBox ctb = new TextBox { Width = 48, Height = 18, FontSize = 8.5, TextAlignment = TextAlignment.Center, VerticalContentAlignment = VerticalAlignment.Center, BorderThickness = new Thickness(1), BorderBrush = new SolidColorBrush(COL_ACCENT), Padding = new Thickness(1, 0, 1, 0) };
                ctb.LostFocus += DiaLostFocus;
                csp.Children.Add(ctb);
                csp.Children.Add(new TextBlock { Text = "mm/\"", FontSize = 7, Foreground = new SolidColorBrush(COL_SUB), HorizontalAlignment = HorizontalAlignment.Center });
                Button cbtn = new Button { Content = "▶", FontSize = 8, Width = 22, Height = 16, Margin = new Thickness(0, 2, 0, 0), Cursor = Cursors.Hand, Background = new SolidColorBrush(COL_ACCENT), Foreground = Brushes.White, BorderThickness = new Thickness(0) };
                cbtn.Click += (s, e) => { if (ParseDia(ctb.Text, out double v)) for (int ri = 0; ri < _rows; ri++) _txtDiameter[ri, colIdx].Text = v.ToString("F1"); };
                csp.Children.Add(cbtn);
                Grid.SetRow(csp, todosYRow); Grid.SetColumn(csp, HoleCol(c));
                g.Children.Add(csp);
            }

            // ── MARCO (inside schematic grid, bottom row) ──
            g.RowDefinitions.Add(new RowDefinition { Height = new GridLength(40) });
            int marcoRow = g.RowDefinitions.Count - 1;
            WrapPanel mRow = new WrapPanel { VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0, 0, 0) };
            mRow.Children.Add(new TextBlock { Text = "Marco:", FontSize = 12, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(COL_ACCENT), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 10, 0) });
            mRow.Children.Add(MkField("Ancho", out _txtAncho, ResultData.Ancho));
            mRow.Children.Add(MkField("Altura", out _txtAltura, ResultData.Altura));
            Button btnCalcMarco = ABtn("\u03A3 Calcular", COL_ACCENT, Colors.White, 100);
            btnCalcMarco.ToolTip = "Ancho = \u03A3 Dist X  |  Altura = \u03A3 Dist Y";
            btnCalcMarco.Click += (s, e) =>
            {
                double sumX = 0;
                for (int i = 0; i < _dxCount; i++) { if (double.TryParse(_txtDistX[i].Text, out double v)) sumX += v; }
                _txtAncho.Text = sumX.ToString("F1");
                double sumY = 0;
                for (int i = 0; i < _dyCount; i++) { if (double.TryParse(_txtDistY[i].Text, out double v)) sumY += v; }
                _txtAltura.Text = sumY.ToString("F1");
            };
            mRow.Children.Add(btnCalcMarco);
            Grid.SetRow(mRow, marcoRow); Grid.SetColumn(mRow, 0);
            Grid.SetColumnSpan(mRow, g.ColumnDefinitions.Count);
            g.Children.Add(mRow);

            // ── BULK TOOLS (inside schematic grid) ──
            g.RowDefinitions.Add(new RowDefinition { Height = new GridLength(40) });
            int bulkRow = g.RowDefinitions.Count - 1;
            WrapPanel bulkWrap = BuildBulk();
            bulkWrap.Margin = new Thickness(4, 0, 0, 0);
            Grid.SetRow(bulkWrap, bulkRow); Grid.SetColumn(bulkWrap, 0);
            Grid.SetColumnSpan(bulkWrap, g.ColumnDefinitions.Count);
            g.Children.Add(bulkWrap);

            schemCard.Child = g;
            root.Children.Add(schemCard);

            root.Children.Add(BuildActions());
        }

        // ═══════════════════════════════════════════
        void UpdHole(int r, int c, bool vis)
        {
            bool orig = _original.Visible[r, c];
            bool changed = vis != orig;
            if (vis) { _circles[r, c].Fill = changed ? BR_PEND : BR_ON; _circles[r, c].Stroke = changed ? BR_STR_PEND : BR_STR_ON; }
            else { _circles[r, c].Fill = changed ? new SolidColorBrush(Color.FromArgb(60, 255, 193, 7)) : BR_OFF; _circles[r, c].Stroke = changed ? BR_STR_PEND : BR_STR_OFF; }
            _txtDiameter[r, c].IsEnabled = vis;
            _txtDiameter[r, c].Background = new SolidColorBrush(Color.FromArgb(vis ? (byte)180 : (byte)60, 255, 255, 255));
            _chkVisible[r, c].Foreground = new SolidColorBrush(vis ? Colors.White : Color.FromRgb(100, 100, 100));
        }

        WrapPanel BuildBulk()
        {
            WrapPanel w = new WrapPanel();
            Button on = ABtn("✔ Todos ON", COL_GREEN, Colors.White, 95); on.Click += (s, e) => SetAll(true); w.Children.Add(on);
            Button off = ABtn("✖ Todos OFF", COL_RED, Colors.White, 95); off.Click += (s, e) => SetAll(false); w.Children.Add(off);
            w.Children.Add(new Border { Width = 10 });
            w.Children.Add(SL("Fila:"));
            ComboBox cr = new ComboBox { Width = 42, Height = 24, FontSize = 10, VerticalContentAlignment = VerticalAlignment.Center }; for (int i = 1; i <= _rows; i++) cr.Items.Add(i); cr.SelectedIndex = 0; w.Children.Add(cr);
            Button tr = SBtn("Toggle"); tr.Click += (s, e) => { int ri = cr.SelectedIndex; for (int c = 0; c < _cols; c++) _chkVisible[ri, c].IsChecked = !(_chkVisible[ri, c].IsChecked == true); }; w.Children.Add(tr);
            w.Children.Add(new Border { Width = 6 });
            w.Children.Add(SL("Col:"));
            ComboBox cc = new ComboBox { Width = 42, Height = 24, FontSize = 10, VerticalContentAlignment = VerticalAlignment.Center }; for (int i = 1; i <= _cols; i++) cc.Items.Add(i); cc.SelectedIndex = 0; w.Children.Add(cc);
            Button tc = SBtn("Toggle"); tc.Click += (s, e) => { int ci = cc.SelectedIndex; for (int r = 0; r < _rows; r++) _chkVisible[r, ci].IsChecked = !(_chkVisible[r, ci].IsChecked == true); }; w.Children.Add(tc);
            w.Children.Add(new Border { Width = 10 });
            w.Children.Add(SL("ø marcados:"));
            TextBox tb = SI(44); Border bd = MkIB(); WF(tb, bd); bd.Child = tb; w.Children.Add(bd);
            w.Children.Add(new TextBlock { Text = "mm/\"", FontSize = 8, Foreground = new SolidColorBrush(COL_SUB), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(2, 0, 4, 0) });
            Button ba = SBtn("Aplicar ø"); ba.Click += (s, e) => { if (!ParseDia(tb.Text, out double d)) return; for (int r = 0; r < _rows; r++) for (int c = 0; c < _cols; c++) if (_chkVisible[r, c].IsChecked == true) _txtDiameter[r, c].Text = d.ToString("F1"); }; w.Children.Add(ba);
            return w;
        }

        StackPanel BuildActions()
        {
            StackPanel row = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 2, 0, 0) };
            Button rst = ABtn("Restaurar", COL_BTN, COL_TEXT, 100); rst.Click += (s, e) => { ResultData = _original.Clone(); Reload(); }; row.Children.Add(rst);
            Button cnc = ABtn("Cancelar", COL_BTN, COL_TEXT, 85); cnc.Click += (s, e) => { DialogResult = false; Close(); }; row.Children.Add(cnc);
            Button app = ABtn($"Aplicar a {_instanceCount} inst.", COL_ACCENT, Colors.White, 190); app.FontSize = 13; app.FontWeight = FontWeights.SemiBold; app.Height = 36; app.Click += DoApply; row.Children.Add(app);
            return row;
        }

        void DoApply(object sender, RoutedEventArgs e)
        {
            bool err = false;
            if (double.TryParse(_txtAncho.Text, out double a)) ResultData.Ancho = a; else err = ME(_txtAncho);
            if (double.TryParse(_txtAltura.Text, out double al)) ResultData.Altura = al; else err = ME(_txtAltura);
            for (int i = 0; i < _dxCount; i++) { if (double.TryParse(_txtDistX[i].Text, out double v)) ResultData.DistX[i] = v; else err = ME(_txtDistX[i]); }
            for (int i = 0; i < _dyCount; i++) { if (double.TryParse(_txtDistY[i].Text, out double v)) ResultData.DistY[i] = v; else err = ME(_txtDistY[i]); }
            for (int r = 0; r < _rows; r++) for (int c = 0; c < _cols; c++)
                { ResultData.Visible[r, c] = _chkVisible[r, c].IsChecked == true; if (ParseDia(_txtDiameter[r, c].Text, out double d)) ResultData.Diameter[r, c] = d; else if (_chkVisible[r, c].IsChecked == true) err = ME(_txtDiameter[r, c]); else ResultData.Diameter[r, c] = 0; }
            if (err) { MessageBox.Show("Campos inválidos (bordes rojos).", "Error", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
            DialogResult = true; Close();
        }

        bool ME(TextBox t) { t.BorderBrush = new SolidColorBrush(COL_RED); t.BorderThickness = new Thickness(2); return true; }

        bool ParseDia(string text, out double mm)
        {
            if (text == null) { mm = 0; return false; }
            string s = text.Trim();
            if (s.EndsWith("\""))
            {
                s = s.TrimEnd('"').Trim();
                if (double.TryParse(s, out double inches)) { mm = inches * 25.4; return true; }
                mm = 0; return false;
            }
            return double.TryParse(s, out mm);
        }

        void DiaLostFocus(object sender, RoutedEventArgs e)
        {
            TextBox t = sender as TextBox; if (t == null) return;
            if (ParseDia(t.Text, out double mm)) t.Text = mm.ToString("F1");
        }

        void SetAll(bool st) { for (int r = 0; r < _rows; r++) for (int c = 0; c < _cols; c++) _chkVisible[r, c].IsChecked = st; }

        void Reload()
        {
            _txtAncho.Text = ResultData.Ancho.ToString("F1"); _txtAltura.Text = ResultData.Altura.ToString("F1");
            for (int i = 0; i < _dxCount; i++) _txtDistX[i].Text = ResultData.DistX[i].ToString("F1");
            for (int i = 0; i < _dyCount; i++) _txtDistY[i].Text = ResultData.DistY[i].ToString("F1");
            for (int r = 0; r < _rows; r++) for (int c = 0; c < _cols; c++) { _chkVisible[r, c].IsChecked = ResultData.Visible[r, c]; _txtDiameter[r, c].Text = ResultData.Diameter[r, c].ToString("F1"); UpdHole(r, c, ResultData.Visible[r, c]); }
        }

        // ── UI helpers ──
        StackPanel LegendDot(SolidColorBrush f, string l)
        {
            StackPanel sp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 14, 0) };
            sp.Children.Add(new Ellipse { Width = 12, Height = 12, Fill = f, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 4, 0) });
            sp.Children.Add(new TextBlock { Text = l, FontSize = 10, Foreground = new SolidColorBrush(COL_SUB), VerticalAlignment = VerticalAlignment.Center });
            return sp;
        }

        StackPanel MkField(string lbl, out TextBox txt, double val)
        {
            StackPanel sp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 8, 0) };
            sp.Children.Add(new TextBlock { Text = lbl + ":", FontSize = 11, Foreground = new SolidColorBrush(COL_TEXT), VerticalAlignment = VerticalAlignment.Center, Width = 48 });
            Border b = MkIB();
            txt = new TextBox { Width = 56, Height = 24, FontSize = 11, Text = val.ToString("F1"), TextAlignment = TextAlignment.Center, VerticalContentAlignment = VerticalAlignment.Center, BorderThickness = new Thickness(0), Background = Brushes.Transparent, Padding = new Thickness(2, 0, 2, 0) };
            WF(txt, b); b.Child = txt; sp.Children.Add(b);
            return sp;
        }

        Border MkCard() => new Border { CornerRadius = new CornerRadius(8), BorderBrush = new SolidColorBrush(COL_BORDER), BorderThickness = new Thickness(1), Background = Brushes.White, Padding = new Thickness(10) };
        Border MkIB() => new Border { CornerRadius = new CornerRadius(5), BorderBrush = new SolidColorBrush(COL_BORDER), BorderThickness = new Thickness(1), Background = Brushes.White, Margin = new Thickness(0, 0, 4, 0) };
        void WF(TextBox t, Border b) { t.GotFocus += (s, e) => b.BorderBrush = new SolidColorBrush(COL_ACCENT); t.LostFocus += (s, e) => b.BorderBrush = new SolidColorBrush(COL_BORDER); }
        TextBlock SL(string t) => new TextBlock { Text = t, FontSize = 10, Foreground = new SolidColorBrush(COL_TEXT), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 4, 0) };
        TextBox SI(double w) => new TextBox { Width = w, Height = 22, FontSize = 10, TextAlignment = TextAlignment.Center, VerticalContentAlignment = VerticalAlignment.Center, BorderThickness = new Thickness(0), Background = Brushes.Transparent, Padding = new Thickness(2, 0, 2, 0) };
        Button SBtn(string t) { Button b = new Button { Content = t, Height = 24, FontSize = 10, Padding = new Thickness(8, 0, 8, 0), Margin = new Thickness(0, 0, 4, 0), Cursor = Cursors.Hand, Background = new SolidColorBrush(COL_BTN), Foreground = new SolidColorBrush(COL_TEXT), BorderThickness = new Thickness(0) }; b.Template = RBT(COL_BTN); return b; }
        Button ABtn(string t, Color bg, Color fg, double w) { Button b = new Button { Content = t, Width = w, Height = 32, FontSize = 11, Margin = new Thickness(0, 0, 6, 0), Cursor = Cursors.Hand, Foreground = new SolidColorBrush(fg), Background = new SolidColorBrush(bg), BorderThickness = new Thickness(0) }; b.Template = RBT(bg); return b; }
        ControlTemplate RBT(Color bg) { var t = new ControlTemplate(typeof(Button)); var bd = new FrameworkElementFactory(typeof(Border)); bd.SetValue(Border.CornerRadiusProperty, new CornerRadius(6)); bd.SetValue(Border.BackgroundProperty, new SolidColorBrush(bg)); bd.SetValue(Border.PaddingProperty, new Thickness(10, 4, 10, 4)); var cp2 = new FrameworkElementFactory(typeof(ContentPresenter)); cp2.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center); cp2.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center); bd.AppendChild(cp2); t.VisualTree = bd; return t; }
    }
}