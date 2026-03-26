using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;

using Color = System.Windows.Media.Color;
using TextBox = System.Windows.Controls.TextBox;
using Button = System.Windows.Controls.Button;
using Grid = System.Windows.Controls.Grid;
using Rectangle = System.Windows.Shapes.Rectangle;

namespace HMVTools
{
    public class FoundationLayoutWindow : Window
    {
        // ═══════════════════════════════════════════════
        //  Public result
        // ═══════════════════════════════════════════════
        public List<FoundationItemData> Items { get; private set; }

        // ═══════════════════════════════════════════════
        //  Colors (matching USDTuberia palette)
        // ═══════════════════════════════════════════════
        static readonly Color COL_BG = Color.FromRgb(245, 245, 248);
        static readonly Color COL_ACCENT = Color.FromRgb(0, 120, 212);
        static readonly Color COL_GREEN = Color.FromRgb(40, 167, 69);
        static readonly Color COL_RED = Color.FromRgb(220, 53, 69);
        static readonly Color COL_BORDER = Color.FromRgb(200, 200, 210);
        static readonly Color COL_TEXT = Color.FromRgb(30, 30, 30);
        static readonly Color COL_SUB = Color.FromRgb(120, 120, 130);
        static readonly Color COL_BTN = Color.FromRgb(240, 240, 243);

        // ═══════════════════════════════════════════════
        //  State
        // ═══════════════════════════════════════════════
        private readonly string _linkLabel;
        private readonly double _surveyOffsetM;
        private Canvas _canvas;
        private TextBox[] _txtNtce;
        private Rectangle[] _canvasRects;
        private Border[] _rowBorders;
        private int _selectedIndex = -1;

        // ═══════════════════════════════════════════════
        //  Constructor
        // ═══════════════════════════════════════════════
        public FoundationLayoutWindow(
            List<FoundationItemData> items,
            string linkLabel,
            double surveyOffsetM)
        {
            Items = items;
            _linkLabel = linkLabel;
            _surveyOffsetM = surveyOffsetM;
            _txtNtce = new TextBox[items.Count];
            _canvasRects = new Rectangle[items.Count];
            _rowBorders = new Border[items.Count];

            BuildUI();
        }

        // ═══════════════════════════════════════════════
        //  Build UI
        // ═══════════════════════════════════════════════
        void BuildUI()
        {
            Title = "HMV Tools — Foundation Layout & Elevation Manager";
            Width = 960; Height = 720;
            MinWidth = 800; MinHeight = 600;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Background = new SolidColorBrush(COL_BG);
            ResizeMode = ResizeMode.CanResize;

            // Outer scroll for the whole window
            ScrollViewer outerScroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Padding = new Thickness(20)
            };
            Content = outerScroll;

            StackPanel root = new StackPanel();
            outerScroll.Content = root;

            // ── Header ──────────────────────────────────────
            root.Children.Add(new TextBlock
            {
                Text = "Foundation Layout & Elevation Manager",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(COL_TEXT),
                Margin = new Thickness(0, 0, 0, 2)
            });

            int napCount = Items.Count(i => i.HasNap);
            int noNapCount = Items.Count - napCount;
            string sub = $"{Items.Count} foundation(s)  ·  Link: {_linkLabel}  ·  Survey offset: {_surveyOffsetM:F3} m";
            if (noNapCount > 0)
                sub += $"  ·  {noNapCount} without NAP (disabled)";

            root.Children.Add(new TextBlock
            {
                Text = sub,
                FontSize = 11,
                Foreground = new SolidColorBrush(COL_SUB),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 12)
            });

            // ── Canvas card ─────────────────────────────────
            Border canvasCard = MkCard();
            canvasCard.MinHeight = 200;
            canvasCard.Margin = new Thickness(0, 0, 0, 12);

            _canvas = new Canvas
            {
                ClipToBounds = true,
                MinHeight = 180,
                Height = 200,
                Background = Brushes.Transparent
            };
            _canvas.Loaded += (s, e) => DrawCanvas();
            _canvas.SizeChanged += (s, e) => DrawCanvas();

            canvasCard.Child = _canvas;
            root.Children.Add(canvasCard);

            // ── Legend ───────────────────────────────────────
            WrapPanel legend = new WrapPanel
            {
                Margin = new Thickness(0, 0, 0, 8)
            };
            legend.Children.Add(LegendRect(
                new SolidColorBrush(Color.FromArgb(60, 220, 53, 69)),
                new SolidColorBrush(COL_RED), "Foundation"));
            legend.Children.Add(LegendRect(
                new SolidColorBrush(Color.FromArgb(100, 0, 120, 212)),
                new SolidColorBrush(COL_ACCENT), "Selected"));
            root.Children.Add(legend);

            // ── Data table card ─────────────────────────────
            Border tableCard = MkCard();
            tableCard.Margin = new Thickness(0, 0, 0, 12);

            StackPanel tablePanel = new StackPanel();
            tableCard.Child = tablePanel;

            // Header row
            Grid hdr = MakeHeaderRow();
            tablePanel.Children.Add(hdr);

            // Separator
            tablePanel.Children.Add(new Border
            {
                Height = 1,
                Background = new SolidColorBrush(COL_BORDER),
                Margin = new Thickness(0, 4, 0, 4)
            });

            // Scrollable data rows
            ScrollViewer dataScroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                MaxHeight = 320
            };
            StackPanel dataPanel = new StackPanel();
            dataScroll.Content = dataPanel;

            for (int i = 0; i < Items.Count; i++)
                dataPanel.Children.Add(MakeDataRow(i));

            tablePanel.Children.Add(dataScroll);
            root.Children.Add(tableCard);

            // ── Apply button row ────────────────────────────
            root.Children.Add(BuildActions());
        }

        // ═══════════════════════════════════════════════
        //  Canvas drawing
        // ═══════════════════════════════════════════════
        void DrawCanvas()
        {
            _canvas.Children.Clear();
            double cw = _canvas.ActualWidth;
            double ch = _canvas.ActualHeight;
            if (cw < 20 || ch < 20 || Items.Count == 0) return;

            // Global bounds (meters)
            double gMinX = Items.Min(i => i.PlanMinXm);
            double gMinY = Items.Min(i => i.PlanMinYm);
            double gMaxX = Items.Max(i => i.PlanMaxXm);
            double gMaxY = Items.Max(i => i.PlanMaxYm);

            double rangeX = gMaxX - gMinX;
            double rangeY = gMaxY - gMinY;

            // Ensure minimum range to avoid division by zero
            if (rangeX < 0.5) { double cx = (gMinX + gMaxX) / 2; gMinX = cx - 1; rangeX = 2; }
            if (rangeY < 0.5) { double cy = (gMinY + gMaxY) / 2; gMinY = cy - 1; rangeY = 2; }

            // Include foundation sizes in range
            double maxFndW = Items.Max(i => i.PlanMaxXm - i.PlanMinXm);
            double maxFndH = Items.Max(i => i.PlanMaxYm - i.PlanMinYm);
            rangeX = Math.Max(rangeX, maxFndW * 1.5);
            rangeY = Math.Max(rangeY, maxFndH * 1.5);

            double pad = 35;
            double usableW = cw - 2 * pad;
            double usableH = ch - 2 * pad;
            double scale = Math.Min(usableW / rangeX, usableH / rangeY);

            // Center offset
            double drawW = rangeX * scale;
            double drawH = rangeY * scale;
            double offX = pad + (usableW - drawW) / 2;
            double offY = pad + (usableH - drawH) / 2;

            // Origin label (0,0) at bottom-left
            TextBlock originLbl = new TextBlock
            {
                Text = "(0, 0)",
                FontSize = 9,
                Foreground = new SolidColorBrush(COL_SUB),
                FontStyle = FontStyles.Italic
            };
            Canvas.SetLeft(originLbl, offX - 4);
            Canvas.SetTop(originLbl, ch - offY + 4);
            _canvas.Children.Add(originLbl);

            // Draw foundations
            for (int idx = 0; idx < Items.Count; idx++)
            {
                var item = Items[idx];
                double fndW = (item.PlanMaxXm - item.PlanMinXm) * scale;
                double fndH = (item.PlanMaxYm - item.PlanMinYm) * scale;
                double x = offX + (item.PlanMinXm - gMinX) * scale;
                // Flip Y (canvas top=0, plan bottom=0)
                double y = (ch - offY) - (item.PlanMinYm - gMinY) * scale - fndH;

                bool isSelected = idx == _selectedIndex;

                Rectangle rect = new Rectangle
                {
                    Width = Math.Max(fndW, 6),
                    Height = Math.Max(fndH, 6),
                    Stroke = new SolidColorBrush(isSelected ? COL_ACCENT : COL_RED),
                    StrokeThickness = isSelected ? 3 : 1.5,
                    Fill = new SolidColorBrush(isSelected
                        ? Color.FromArgb(100, 0, 120, 212)
                        : Color.FromArgb(60, 220, 53, 69)),
                    Cursor = Cursors.Hand
                };
                int capturedIdx = idx;
                rect.MouseLeftButtonDown += (s, e) => SelectRow(capturedIdx);

                Canvas.SetLeft(rect, x);
                Canvas.SetTop(rect, y);
                _canvas.Children.Add(rect);
                _canvasRects[idx] = rect;

                // Index label
                TextBlock numLbl = new TextBlock
                {
                    Text = (idx + 1).ToString(),
                    FontSize = 10,
                    FontWeight = FontWeights.Bold,
                    Foreground = new SolidColorBrush(isSelected ? COL_ACCENT : COL_RED)
                };
                Canvas.SetLeft(numLbl, x + 3);
                Canvas.SetTop(numLbl, y + 2);
                _canvas.Children.Add(numLbl);
            }
        }

        void SelectRow(int idx)
        {
            _selectedIndex = idx;
            DrawCanvas();

            // Highlight row borders
            for (int i = 0; i < _rowBorders.Length; i++)
            {
                if (_rowBorders[i] == null) continue;
                _rowBorders[i].Background = i == idx
                    ? new SolidColorBrush(Color.FromArgb(30, 0, 120, 212))
                    : Brushes.Transparent;
            }
        }

        // ═══════════════════════════════════════════════
        //  Table header
        // ═══════════════════════════════════════════════
        Grid MakeHeaderRow()
        {
            Grid g = new Grid { Margin = new Thickness(0, 0, 0, 2) };
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) });   // #
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Name
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });  // NTCE
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });  // NAP Survey
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });  // NAP Project

            AddHeaderCell(g, 0, "#");
            AddHeaderCell(g, 1, "Foundation");
            AddHeaderCell(g, 2, "NTCE (m survey)");
            AddHeaderCell(g, 3, "NAP Survey (m)");
            AddHeaderCell(g, 4, "NAP Project (m)");

            return g;
        }

        void AddHeaderCell(Grid g, int col, string text)
        {
            TextBlock tb = new TextBlock
            {
                Text = text,
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(COL_ACCENT),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0, 4, 0)
            };
            Grid.SetColumn(tb, col);
            g.Children.Add(tb);
        }

        // ═══════════════════════════════════════════════
        //  Data row
        // ═══════════════════════════════════════════════
        Border MakeDataRow(int idx)
        {
            var item = Items[idx];

            Border rowBorder = new Border
            {
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(0, 3, 0, 3),
                Margin = new Thickness(0, 1, 0, 1),
                Background = Brushes.Transparent,
                Cursor = Cursors.Hand
            };
            _rowBorders[idx] = rowBorder;

            int capturedIdx = idx;
            rowBorder.MouseLeftButtonDown += (s, e) => SelectRow(capturedIdx);

            Grid g = new Grid();
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) });   // #
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Name
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });  // NTCE
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });  // NAP Survey
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });  // NAP Project

            // # column
            TextBlock numTb = new TextBlock
            {
                Text = (idx + 1).ToString(),
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(COL_RED),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0, 4, 0)
            };
            Grid.SetColumn(numTb, 0);
            g.Children.Add(numTb);

            // Name column
            TextBlock nameTb = new TextBlock
            {
                Text = item.Name,
                FontSize = 10.5,
                Foreground = new SolidColorBrush(item.CanEdit ? COL_TEXT : COL_SUB),
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis,
                Margin = new Thickness(4, 0, 4, 0)
            };
            Grid.SetColumn(nameTb, 1);
            g.Children.Add(nameTb);

            // NTCE column (editable)
            Border ntceBorder = MkIB();
            ntceBorder.Width = 100;
            ntceBorder.HorizontalAlignment = HorizontalAlignment.Left;

            TextBox ntceTb = new TextBox
            {
                Width = 94,
                Height = 24,
                FontSize = 11,
                Text = item.NtceSurveyM.ToString("F3"),
                TextAlignment = TextAlignment.Center,
                VerticalContentAlignment = VerticalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(2, 0, 2, 0),
                IsEnabled = item.CanEdit
            };

            if (!item.CanEdit)
            {
                ntceTb.Foreground = new SolidColorBrush(COL_SUB);
                ntceBorder.Background = new SolidColorBrush(
                    Color.FromRgb(235, 235, 238));
                ntceTb.ToolTip = "Disabled — no floor detected (NAP)";
            }

            WF(ntceTb, ntceBorder);
            ntceBorder.Child = ntceTb;
            _txtNtce[idx] = ntceTb;

            Grid.SetColumn(ntceBorder, 2);
            g.Children.Add(ntceBorder);

            // NAP Survey column
            TextBlock napSurvTb = new TextBlock
            {
                Text = item.HasNap ? item.NapSurveyM.ToString("F3") : "—",
                FontSize = 11,
                Foreground = new SolidColorBrush(item.HasNap ? COL_GREEN : COL_SUB),
                FontWeight = item.HasNap ? FontWeights.SemiBold : FontWeights.Normal,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(8, 0, 4, 0)
            };
            Grid.SetColumn(napSurvTb, 3);
            g.Children.Add(napSurvTb);

            // NAP Project column
            TextBlock napProjTb = new TextBlock
            {
                Text = item.HasNap ? item.NapProjectM.ToString("F3") : "—",
                FontSize = 11,
                Foreground = new SolidColorBrush(item.HasNap ? COL_TEXT : COL_SUB),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(8, 0, 4, 0)
            };
            Grid.SetColumn(napProjTb, 4);
            g.Children.Add(napProjTb);

            rowBorder.Child = g;
            return rowBorder;
        }

        // ═══════════════════════════════════════════════
        //  Actions row
        // ═══════════════════════════════════════════════
        StackPanel BuildActions()
        {
            StackPanel row = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 4, 0, 0)
            };

            Button btnCancel = ABtn("Cancel", COL_BTN, COL_TEXT, 100);
            btnCancel.Click += (s, e) => { DialogResult = false; Close(); };
            row.Children.Add(btnCancel);

            int editableCount = Items.Count(i => i.CanEdit);
            Button btnApply = ABtn(
                $"Apply to {editableCount} foundations",
                COL_ACCENT, Colors.White, 210);
            btnApply.FontSize = 13;
            btnApply.FontWeight = FontWeights.SemiBold;
            btnApply.Height = 36;
            btnApply.Click += DoApply;
            row.Children.Add(btnApply);

            return row;
        }

        // ═══════════════════════════════════════════════
        //  Apply handler
        // ═══════════════════════════════════════════════
        void DoApply(object sender, RoutedEventArgs e)
        {
            bool hasError = false;

            for (int i = 0; i < Items.Count; i++)
            {
                if (!Items[i].CanEdit) continue;

                string text = _txtNtce[i].Text.Trim();
                if (double.TryParse(text, out double val))
                {
                    Items[i].NewNtceSurveyM = val;
                    // Reset border
                    Border parent = _txtNtce[i].Parent as Border;
                    if (parent != null)
                        parent.BorderBrush = new SolidColorBrush(COL_BORDER);
                }
                else
                {
                    hasError = true;
                    MarkError(_txtNtce[i]);
                }
            }

            if (hasError)
            {
                MessageBox.Show(
                    "Invalid values detected (highlighted in red). Please fix before applying.",
                    "Validation Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Check if anything changed
            int changes = 0;
            for (int i = 0; i < Items.Count; i++)
            {
                if (!Items[i].CanEdit) continue;
                if (Math.Abs(Items[i].NewNtceSurveyM - Items[i].NtceSurveyM) > 0.0001)
                    changes++;
            }

            if (changes == 0)
            {
                if (MessageBox.Show(
                    "No NTCE values were changed. Close without changes?",
                    "No Changes",
                    MessageBoxButton.YesNo, MessageBoxImage.Question)
                    == MessageBoxResult.Yes)
                {
                    DialogResult = false; Close();
                }
                return;
            }

            DialogResult = true;
            Close();
        }

        void MarkError(TextBox t)
        {
            Border parent = t.Parent as Border;
            if (parent != null)
            {
                parent.BorderBrush = new SolidColorBrush(COL_RED);
                parent.BorderThickness = new Thickness(2);
            }
        }

        // ═══════════════════════════════════════════════
        //  UI helpers (same patterns as USDTuberiaWindow)
        // ═══════════════════════════════════════════════

        Border MkCard() => new Border
        {
            CornerRadius = new CornerRadius(8),
            BorderBrush = new SolidColorBrush(COL_BORDER),
            BorderThickness = new Thickness(1),
            Background = Brushes.White,
            Padding = new Thickness(10)
        };

        Border MkIB() => new Border
        {
            CornerRadius = new CornerRadius(5),
            BorderBrush = new SolidColorBrush(COL_BORDER),
            BorderThickness = new Thickness(1),
            Background = Brushes.White,
            Margin = new Thickness(0, 0, 4, 0)
        };

        void WF(TextBox t, Border b)
        {
            t.GotFocus += (s, e) =>
                b.BorderBrush = new SolidColorBrush(COL_ACCENT);
            t.LostFocus += (s, e) =>
                b.BorderBrush = new SolidColorBrush(COL_BORDER);
        }

        Button ABtn(string t, Color bg, Color fg, double w)
        {
            Button b = new Button
            {
                Content = t,
                Width = w,
                Height = 32,
                FontSize = 11,
                Margin = new Thickness(0, 0, 6, 0),
                Cursor = Cursors.Hand,
                Foreground = new SolidColorBrush(fg),
                Background = new SolidColorBrush(bg),
                BorderThickness = new Thickness(0)
            };
            b.Template = RBT(bg);
            return b;
        }

        ControlTemplate RBT(Color bg)
        {
            var t = new ControlTemplate(typeof(Button));
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
            t.VisualTree = bd;
            return t;
        }

        StackPanel LegendRect(SolidColorBrush fill,
            SolidColorBrush stroke, string label)
        {
            StackPanel sp = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 16, 0)
            };
            sp.Children.Add(new Rectangle
            {
                Width = 14,
                Height = 14,
                Fill = fill,
                Stroke = stroke,
                StrokeThickness = 1.5,
                RadiusX = 2,
                RadiusY = 2,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 5, 0)
            });
            sp.Children.Add(new TextBlock
            {
                Text = label,
                FontSize = 10,
                Foreground = new SolidColorBrush(COL_SUB),
                VerticalAlignment = VerticalAlignment.Center
            });
            return sp;
        }
    }
}