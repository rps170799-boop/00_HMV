using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using ACadSharp;
using ACadSharp.IO;
using AcadText = ACadSharp.Entities.TextEntity;
using AcadMText = ACadSharp.Entities.MText;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class DwgConvertCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View view = doc.ActiveView;

            List<ImportInstance> dwgs = new FilteredElementCollector(doc)
                .OfClass(typeof(ImportInstance))
                .Cast<ImportInstance>()
                .ToList();

            if (dwgs.Count == 0)
            {
                TaskDialog.Show("HMV Tools", "No DWG imports found.");
                return Result.Cancelled;
            }

            List<string> names = new List<string>();
            foreach (var d in dwgs)
            {
                string name = d.Category != null ? d.Category.Name : "Unknown";
                names.Add(name + " [ID: " + d.Id.IntegerValue + "]");
            }

            DwgConvertWindow win = new DwgConvertWindow(names);
            if (win.ShowDialog() != true)
                return Result.Cancelled;

            List<ImportInstance> selectedDwgs = new List<ImportInstance>();
            foreach (int idx in win.SelectedIndices)
                selectedDwgs.Add(dwgs[idx]);

            string font = win.SelectedFont;

            if (win.Action == DwgConvertAction.ConvertLines)
                return ConvertLines(doc, view, selectedDwgs);
            else if (win.Action == DwgConvertAction.StandardizeTexts)
                return StandardizeTexts(doc, view, selectedDwgs, font);

            return Result.Cancelled;
        }

        // ========== LINES ==========
        private Result ConvertLines(Document doc, View view,
            List<ImportInstance> selectedDwgs)
        {
            int count = 0;
            Category linesCat = doc.Settings.Categories
                .get_Item(BuiltInCategory.OST_Lines);
            var lineStyleCache = new Dictionary<string, GraphicsStyle>();

            using (Transaction t = new Transaction(doc, "DWG to Standardized Lines"))
            {
                t.Start();

                foreach (Category subCat in linesCat.SubCategories)
                {
                    if (subCat.Name.StartsWith("HMV_LINEA"))
                    {
                        GraphicsStyle gs = subCat.GetGraphicsStyle(
                            GraphicsStyleType.Projection);
                        if (gs != null)
                            lineStyleCache[subCat.Name] = gs;
                    }
                }

                foreach (ImportInstance selected in selectedDwgs)
                {
                    Options geoOptions = new Options();
                    geoOptions.View = view;
                    geoOptions.ComputeReferences = true;
                    GeometryElement geoElem = selected.get_Geometry(geoOptions);

                    var curveData = new List<(Curve curve, GraphicsStyle style)>();
                    GetCurves(doc, geoElem, curveData);

                    foreach (var item in curveData)
                    {
                        try
                        {
                            if (item.curve.Length < 0.003) continue;

                            DetailCurve dc = doc.Create.NewDetailCurve(
                                view, item.curve);
                            string styleName = GetHmvLineStyleName(
                                doc, item.style);

                            if (styleName != null)
                            {
                                if (!lineStyleCache.TryGetValue(styleName,
                                    out GraphicsStyle hmvStyle))
                                {
                                    hmvStyle = CreateLineStyle(doc, linesCat,
                                        item.style, styleName);
                                    if (hmvStyle != null)
                                        lineStyleCache[styleName] = hmvStyle;
                                }
                                if (hmvStyle != null)
                                    dc.LineStyle = hmvStyle;
                            }
                            count++;
                        }
                        catch { }
                    }
                }
                t.Commit();
            }

            TaskDialog.Show("HMV Tools",
                count + " detail lines created.\n"
                + lineStyleCache.Count + " HMV line styles used.\n\n"
                + "Next: Select DWG(s) again and click 'Standardize Texts'.");
            return Result.Succeeded;
        }

        // ========== TEXTS ==========
        private static readonly double[] AllowedSizes =
            { 1.5, 2.0, 2.5, 3.0, 3.5 };

        private Result StandardizeTexts(Document doc, View view,
            List<ImportInstance> selectedDwgs, string font)
        {
            int textCount = 0;
            int errorCount = 0;
            string lastError = "";
            int purgedStyles = 0;
            int purgedTextTypes = 0;
            var typeCache = new Dictionary<string, TextNoteType>();

            List<DwgTextData> allTexts = new List<DwgTextData>();

            foreach (ImportInstance dwg in selectedDwgs)
            {
                string dwgPath = GetDwgFilePath(doc, dwg);

                if (dwgPath == null || !System.IO.File.Exists(dwgPath))
                {
                    TaskDialog.Show("HMV Tools",
                        "Cannot find DWG file on disk.\n"
                        + "Make sure the DWG is linked (not imported) "
                        + "or browse to the original file.");
                    continue;
                }

                // ============================================================
                // STEP 1: Get the GeometryInstance from Revit
                // This gives us:
                //   - gi.Transform: the REAL transform Revit uses (handles
                //     UTM offset, rotation, unit conversion — everything)
                //   - gi.GetSymbolGeometry(): curves in the DWG's "symbol
                //     space" (the coordinate system BEFORE the transform)
                // ============================================================
                Options geoOpts = new Options();
                geoOpts.View = view;
                geoOpts.ComputeReferences = true;
                GeometryElement geoElem = dwg.get_Geometry(geoOpts);

                Transform giTransform = null;
                List<XYZ> symbolPoints = new List<XYZ>();

                foreach (GeometryObject geoObj in geoElem)
                {
                    if (geoObj is GeometryInstance gi)
                    {
                        giTransform = gi.Transform;

                        // Collect line endpoints from symbol geometry
                        // These are in the DWG's internal symbol space
                        CollectSymbolPoints(
                            gi.GetSymbolGeometry(), symbolPoints);
                        break; // Use first GeometryInstance found
                    }
                }

                if (giTransform == null)
                {
                    TaskDialog.Show("HMV Tools",
                        "Cannot find GeometryInstance in DWG.\n"
                        + "The DWG may not contain visible geometry.");
                    continue;
                }

                // ============================================================
                // STEP 2: Read the DWG file with ACadSharp
                // Get text positions AND line endpoints (for calibration)
                // ============================================================
                try
                {
                    CadDocument cadDoc;
                    using (DwgReader reader = new DwgReader(dwgPath))
                    {
                        cadDoc = reader.Read();
                    }

                    double unitsToMm = GetDwgUnitsToMmFactor(cadDoc);

                    // Collect texts
                    List<DwgTextData> rawTexts = new List<DwgTextData>();

                    foreach (var entity in cadDoc.Entities)
                    {
                        CollectRawText(entity, null, rawTexts);
                    }

                    foreach (var entity in cadDoc.Entities)
                    {
                        if (entity is ACadSharp.Entities.Insert insert
                            && insert.Block != null)
                        {
                            foreach (var be in insert.Block.Entities)
                            {
                                CollectRawText(be, insert, rawTexts);
                            }
                        }
                    }

                    if (rawTexts.Count == 0) continue;

                    // Collect line endpoints from ACadSharp (for calibration)
                    List<double[]> acadPoints = new List<double[]>();
                    foreach (var entity in cadDoc.Entities)
                    {
                        CollectAcadLinePoints(entity, acadPoints);
                    }

                    // ========================================================
                    // STEP 3: Derive the mapping from ACadSharp → symbol space
                    //
                    // Both sources read the same DWG file, so their bounding
                    // boxes correspond to the same physical geometry. The
                    // relationship is a linear mapping:
                    //   symbol_coord = acad_coord * scale + offset
                    //
                    // We compute scale and offset by comparing bounding boxes.
                    // ========================================================

                    // Symbol space bounding box (from Revit)
                    double sMinX = double.MaxValue, sMinY = double.MaxValue;
                    double sMaxX = double.MinValue, sMaxY = double.MinValue;
                    foreach (XYZ sp in symbolPoints)
                    {
                        if (sp.X < sMinX) sMinX = sp.X;
                        if (sp.Y < sMinY) sMinY = sp.Y;
                        if (sp.X > sMaxX) sMaxX = sp.X;
                        if (sp.Y > sMaxY) sMaxY = sp.Y;
                    }

                    // ACadSharp bounding box (from file)
                    double aMinX = double.MaxValue, aMinY = double.MaxValue;
                    double aMaxX = double.MinValue, aMaxY = double.MinValue;
                    foreach (double[] ap in acadPoints)
                    {
                        if (ap[0] < aMinX) aMinX = ap[0];
                        if (ap[1] < aMinY) aMinY = ap[1];
                        if (ap[0] > aMaxX) aMaxX = ap[0];
                        if (ap[1] > aMaxY) aMaxY = ap[1];
                    }

                    double aDW = aMaxX - aMinX;
                    double aDH = aMaxY - aMinY;
                    double sDW = sMaxX - sMinX;
                    double sDH = sMaxY - sMinY;

                    if (aDW < 0.001 || aDH < 0.001
                        || sDW < 0.001 || sDH < 0.001)
                    {
                        TaskDialog.Show("HMV Tools",
                            "Insufficient geometry for calibration.\n"
                            + "ACad points: " + acadPoints.Count + "\n"
                            + "Symbol points: " + symbolPoints.Count);
                        continue;
                    }

                    // Scale and offset: symbol = acad * scale + offset
                    double scaleX = sDW / aDW;
                    double scaleY = sDH / aDH;
                    double offsetX = sMinX - aMinX * scaleX;
                    double offsetY = sMinY - aMinY * scaleY;

                    // Use average scale for uniform mapping
                    // (should be nearly identical for X and Y)
                    double scale = (scaleX + scaleY) / 2.0;

                    // Rotation angle from the GeometryInstance transform
                    double importAngle = Math.Atan2(
                        giTransform.BasisX.Y,
                        giTransform.BasisX.X);

                    // ===== DEBUG DIALOG =====
                    string firstText = rawTexts[0].Text;
                    if (firstText.Length > 30)
                        firstText = firstText.Substring(0, 30) + "...";

                    // Test mapping of first text
                    double testSymX = rawTexts[0].RawX * scaleX + offsetX;
                    double testSymY = rawTexts[0].RawY * scaleY + offsetY;
                    XYZ testRevit = giTransform.OfPoint(
                        new XYZ(testSymX, testSymY, 0));

                    string dbg =
                        "Texts: " + rawTexts.Count + "\n"
                        + "Symbol pts: " + symbolPoints.Count
                        + " | ACad pts: " + acadPoints.Count + "\n\n"
                        + "== BOUNDING BOXES ==\n"
                        + "Symbol: (" + sMinX.ToString("F2")
                        + "," + sMinY.ToString("F2") + ") to ("
                        + sMaxX.ToString("F2") + "," + sMaxY.ToString("F2")
                        + ") size " + sDW.ToString("F2") + "x"
                        + sDH.ToString("F2") + "\n"
                        + "ACad:   (" + aMinX.ToString("F2")
                        + "," + aMinY.ToString("F2") + ") to ("
                        + aMaxX.ToString("F2") + "," + aMaxY.ToString("F2")
                        + ") size " + aDW.ToString("F2") + "x"
                        + aDH.ToString("F2") + "\n\n"
                        + "== CALIBRATION ==\n"
                        + "ScaleX: " + scaleX.ToString("F6")
                        + " | ScaleY: " + scaleY.ToString("F6") + "\n"
                        + "OffsetX: " + offsetX.ToString("F4")
                        + " | OffsetY: " + offsetY.ToString("F4") + "\n\n"
                        + "== GI TRANSFORM ==\n"
                        + "Origin: (" + giTransform.Origin.X.ToString("F4")
                        + ", " + giTransform.Origin.Y.ToString("F4") + ")\n"
                        + "BasisX len: "
                        + giTransform.BasisX.GetLength().ToString("F6") + "\n"
                        + "Rotation: "
                        + (importAngle * 180.0 / Math.PI).ToString("F2")
                        + "°\n\n"
                        + "== FIRST TEXT ==\n"
                        + "\"" + firstText + "\"\n"
                        + "Raw: (" + rawTexts[0].RawX.ToString("F2")
                        + ", " + rawTexts[0].RawY.ToString("F2") + ")\n"
                        + "Symbol: (" + testSymX.ToString("F4")
                        + ", " + testSymY.ToString("F4") + ")\n"
                        + "Revit: (" + testRevit.X.ToString("F4")
                        + ", " + testRevit.Y.ToString("F4") + ")\n"
                        + "Height mm: "
                        + (rawTexts[0].RawHeight * unitsToMm).ToString("F2");

                    TaskDialog.Show("DEBUG - GI Calibration", dbg);

                    // ========================================================
                    // STEP 4: Map all texts through the derived transform
                    //   ACadSharp coord → symbol space → gi.Transform → Revit
                    // ========================================================
                    foreach (var rt in rawTexts)
                    {
                        // ACadSharp → symbol space
                        double symX = rt.RawX * scaleX + offsetX;
                        double symY = rt.RawY * scaleY + offsetY;

                        // Symbol space → Revit world via gi.Transform
                        XYZ revitPt = giTransform.OfPoint(
                            new XYZ(symX, symY, 0));
                        XYZ finalPos = new XYZ(
                            revitPt.X, revitPt.Y, 0);

                        double heightMm = rt.RawHeight * unitsToMm;

                        // Text rotation = own rotation + import rotation
                        double totalRot = rt.RawRotation + importAngle;

                        allTexts.Add(new DwgTextData(
                            rt.Text, finalPos, heightMm, totalRot));
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("HMV Tools",
                        "Error reading DWG:\n" + ex.Message
                        + "\n\n" + ex.StackTrace);
                    continue;
                }
            }

            if (allTexts.Count == 0)
            {
                TaskDialog.Show("HMV Tools",
                    "No texts found in selected DWG(s).");
                return Result.Cancelled;
            }

            using (Transaction t = new Transaction(doc,
                "DWG Texts to Standardized Notes"))
            {
                t.Start();

                // Load existing HMV text types
                foreach (TextNoteType tnt in new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>())
                {
                    if (tnt.Name.StartsWith("HMV_General_"))
                        typeCache[tnt.Name] = tnt;
                }

                TextNoteType baseType = new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .FirstOrDefault();

                if (baseType == null)
                {
                    t.RollBack();
                    TaskDialog.Show("HMV Tools", "No text types found.");
                    return Result.Cancelled;
                }

                foreach (DwgTextData td in allTexts)
                {
                    try
                    {
                        double snapped = SnapToAllowed(td.HeightMm);
                        string typeName = "HMV_General_"
                            + snapped.ToString("0.0") + " " + font;

                        if (!typeCache.TryGetValue(typeName,
                            out TextNoteType hmvType))
                        {
                            hmvType = CreateTextType(
                                doc, baseType, typeName, snapped, font);
                            if (hmvType != null)
                                typeCache[typeName] = hmvType;
                        }

                        if (hmvType != null)
                        {
                            TextNote note = TextNote.Create(
                                doc, view.Id, td.Position,
                                td.Text, hmvType.Id);

                            if (Math.Abs(td.Rotation) > 0.001)
                            {
                                Autodesk.Revit.DB.Line rotAxis =
                                    Autodesk.Revit.DB.Line.CreateUnbound(
                                        td.Position, XYZ.BasisZ);
                                ElementTransformUtils.RotateElement(
                                    doc, note.Id, rotAxis, td.Rotation);
                            }

                            textCount++;
                        }
                        else
                        {
                            errorCount++;
                            lastError = "Type creation failed: " + typeName;
                        }
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        lastError = ex.Message;
                    }
                }

                // ===== PURGE UNUSED LINE STYLES =====
                Category linesCat = doc.Settings.Categories
                    .get_Item(BuiltInCategory.OST_Lines);

                HashSet<ElementId> usedStyleIds = new HashSet<ElementId>();
                foreach (CurveElement ce in new FilteredElementCollector(doc)
                    .OfClass(typeof(CurveElement))
                    .Cast<CurveElement>())
                {
                    try
                    {
                        GraphicsStyle gs = ce.LineStyle as GraphicsStyle;
                        if (gs != null) usedStyleIds.Add(gs.Id);
                    }
                    catch { }
                }

                List<Category> toRemove = new List<Category>();
                List<Category> allSubs = new List<Category>();
                foreach (Category sub in linesCat.SubCategories)
                    allSubs.Add(sub);
                foreach (Category sub in allSubs)
                {
                    if (sub.Name.StartsWith("HMV_LINEA")) continue;
                    if (IsStandardRevitLineStyle(sub.Name)) continue;

                    GraphicsStyle gs = sub.GetGraphicsStyle(
                        GraphicsStyleType.Projection);
                    if (gs == null) continue;

                    if (!usedStyleIds.Contains(gs.Id))
                        toRemove.Add(sub);
                }

                foreach (Category cat in toRemove)
                {
                    try { doc.Delete(cat.Id); purgedStyles++; }
                    catch { }
                }

                // ===== PURGE UNUSED TEXT TYPES =====
                HashSet<ElementId> usedTextTypeIds = new HashSet<ElementId>(
                    new FilteredElementCollector(doc)
                        .OfClass(typeof(TextNote))
                        .Cast<TextNote>()
                        .Select(tn => tn.GetTypeId()));

                List<TextNoteType> allTextTypes =
                    new FilteredElementCollector(doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .ToList();

                foreach (TextNoteType tnt in allTextTypes)
                {
                    if (tnt.Name.StartsWith("HMV_General_")) continue;
                    if (usedTextTypeIds.Contains(tnt.Id)) continue;

                    try { doc.Delete(tnt.Id); purgedTextTypes++; }
                    catch { }
                }

                t.Commit();
            }

            string resultMsg = "Texts created: " + textCount + "\n"
                + "Font: " + font + "\n"
                + "Text types: " + typeCache.Count + " HMV types\n";

            if (errorCount > 0)
            {
                resultMsg += "\nERRORS: " + errorCount + " texts failed\n"
                    + "Last error: " + lastError + "\n";
            }

            resultMsg += "\nCleanup:\n"
                + "Unused line styles purged: " + purgedStyles + "\n"
                + "Unused text types purged: " + purgedTextTypes;

            TaskDialog.Show("HMV Tools", resultMsg);
            return Result.Succeeded;
        }



        private void CollectSymbolPoints(GeometryElement geoElem,
    List<XYZ> points)
        {
            foreach (GeometryObject geoObj in geoElem)
            {
                if (geoObj is Curve curve)
                {
                    points.Add(curve.GetEndPoint(0));
                    points.Add(curve.GetEndPoint(1));
                }
                else if (geoObj is PolyLine poly)
                {
                    foreach (XYZ pt in poly.GetCoordinates())
                        points.Add(pt);
                }
                else if (geoObj is GeometryInstance nestedGi)
                {
                    // GetInstanceGeometry() returns nested block
                    // geometry already in the parent's symbol space
                    CollectSymbolPoints(
                        nestedGi.GetInstanceGeometry(), points);
                }
                else if (geoObj is GeometryElement nestedElem)
                {
                    CollectSymbolPoints(nestedElem, points);
                }
            }
        }



        // ===== ACadSharp TEXT EXTRACTION =====

        private void CollectRawText(ACadSharp.Entities.Entity entity,
            ACadSharp.Entities.Insert parentInsert,
            List<DwgTextData> rawTexts)
        {
            string text = null;
            double x = 0, y = 0, height = 0, rotation = 0;

            if (entity is AcadText textEnt)
            {
                text = textEnt.Value;

                if (textEnt.HorizontalAlignment !=
                        ACadSharp.Entities.TextHorizontalAlignment.Left ||
                    textEnt.VerticalAlignment !=
                        ACadSharp.Entities.TextVerticalAlignmentType.Baseline)
                {
                    x = textEnt.AlignmentPoint.X;
                    y = textEnt.AlignmentPoint.Y;

                    if (Math.Abs(x) < 0.0001 && Math.Abs(y) < 0.0001)
                    {
                        x = textEnt.InsertPoint.X;
                        y = textEnt.InsertPoint.Y;
                    }
                }
                else
                {
                    x = textEnt.InsertPoint.X;
                    y = textEnt.InsertPoint.Y;
                }
                height = textEnt.Height;
                rotation = textEnt.Rotation * Math.PI / 180.0;
            }
            else if (entity is AcadMText mtext)
            {
                text = mtext.Value;
                text = StripMTextFormatting(text);
                x = mtext.InsertPoint.X;
                y = mtext.InsertPoint.Y;
                height = mtext.Height;
                rotation = mtext.Rotation;
            }

            if (string.IsNullOrWhiteSpace(text)) return;

            if (parentInsert != null)
            {
                double ix = parentInsert.InsertPoint.X;
                double iy = parentInsert.InsertPoint.Y;
                double sx = parentInsert.XScale;
                double sy = parentInsert.YScale;
                double insRot = parentInsert.Rotation * Math.PI / 180.0;

                double rx = x * sx * Math.Cos(insRot)
                          - y * sy * Math.Sin(insRot) + ix;
                double ry = x * sx * Math.Sin(insRot)
                          + y * sy * Math.Cos(insRot) + iy;

                x = rx;
                y = ry;
                height *= Math.Abs(sx);
                rotation += insRot;
            }

            var td = new DwgTextData(text, XYZ.Zero, 0, 0);
            td.RawX = x;
            td.RawY = y;
            td.RawHeight = height;
            td.RawRotation = rotation;
            rawTexts.Add(td);
        }

        // ===== ACadSharp LINE POINT EXTRACTION (for calibration) =====

        private void CollectAcadLinePoints(
            ACadSharp.Entities.Entity entity,
            List<double[]> points)
        {
            if (entity is ACadSharp.Entities.Line ln)
            {
                points.Add(new double[] { ln.StartPoint.X, ln.StartPoint.Y });
                points.Add(new double[] { ln.EndPoint.X, ln.EndPoint.Y });
            }
            else if (entity is ACadSharp.Entities.LwPolyline lw)
            {
                foreach (var v in lw.Vertices)
                    points.Add(new double[] { v.Location.X, v.Location.Y });
            }
            else if (entity is ACadSharp.Entities.Polyline2D p2)
            {
                foreach (var v in p2.Vertices)
                    points.Add(new double[] { v.Location.X, v.Location.Y });
            }
            else if (entity is ACadSharp.Entities.Circle circ)
            {
                points.Add(new double[] {
                    circ.Center.X - circ.Radius,
                    circ.Center.Y - circ.Radius });
                points.Add(new double[] {
                    circ.Center.X + circ.Radius,
                    circ.Center.Y + circ.Radius });
            }
            else if (entity is ACadSharp.Entities.Arc arc)
            {
                points.Add(new double[] {
                    arc.Center.X - arc.Radius,
                    arc.Center.Y - arc.Radius });
                points.Add(new double[] {
                    arc.Center.X + arc.Radius,
                    arc.Center.Y + arc.Radius });
            }
            else if (entity is ACadSharp.Entities.Insert ins)
            {
                if (ins.Block != null)
                {
                    double ix = ins.InsertPoint.X;
                    double iy = ins.InsertPoint.Y;
                    double sx = ins.XScale;
                    double sy = ins.YScale;
                    double insRot = ins.Rotation * Math.PI / 180.0;

                    foreach (var be in ins.Block.Entities)
                    {
                        var subPts = new List<double[]>();
                        CollectAcadLinePoints(be, subPts);
                        foreach (var sp in subPts)
                        {
                            double rx = sp[0] * sx * Math.Cos(insRot)
                                      - sp[1] * sy * Math.Sin(insRot) + ix;
                            double ry = sp[0] * sx * Math.Sin(insRot)
                                      + sp[1] * sy * Math.Cos(insRot) + iy;
                            points.Add(new double[] { rx, ry });
                        }
                    }
                }
            }
        }

        // ===== UNIT HELPERS =====

        private double GetDwgUnitsToMmFactor(CadDocument cadDoc)
        {
            try
            {
                int u = (int)cadDoc.Header.InsUnits;
                switch (u)
                {
                    case 1: return 25.4;       // Inches
                    case 2: return 304.8;       // Feet
                    case 3: return 1609344.0;   // Miles
                    case 4: return 1.0;         // Millimeters
                    case 5: return 10.0;        // Centimeters
                    case 6: return 1000.0;      // Meters
                    case 10: return 914.4;       // Yards
                    case 14: return 100.0;       // Decimeters
                    default: return 1.0;          // Assume mm
                }
            }
            catch { return 1.0; }
        }

        // ===== HELPERS =====

        private string StripMTextFormatting(string mtext)
        {
            if (string.IsNullOrEmpty(mtext)) return mtext;

            string result = mtext;
            result = System.Text.RegularExpressions.Regex.Replace(
                result, @"\{\\[^;]+;([^}]*)}", "$1");
            result = System.Text.RegularExpressions.Regex.Replace(
                result, @"\\A\d;", "");
            result = result.Replace("\\P", "\n");
            result = System.Text.RegularExpressions.Regex.Replace(
                result, @"\\[a-zA-Z][^;]*;", "");
            result = result.Replace("{", "").Replace("}", "");

            return result.Trim();
        }

        private string GetDwgFilePath(Document doc, ImportInstance import)
        {
            try
            {
                ElementId typeId = import.GetTypeId();
                Element typeElem = doc.GetElement(typeId);
                if (typeElem != null)
                {
                    ExternalFileReference extRef =
                        typeElem.GetExternalFileReference();
                    if (extRef != null)
                    {
                        ModelPath mp = extRef.GetPath();
                        string path = ModelPathUtils
                            .ConvertModelPathToUserVisiblePath(mp);
                        if (!string.IsNullOrEmpty(path)
                            && System.IO.File.Exists(path))
                            return path;
                    }
                }
            }
            catch { }

            try
            {
                Parameter p = import.get_Parameter(
                    BuiltInParameter.IMPORT_SYMBOL_NAME);
                if (p != null)
                {
                    string name = p.AsString();
                    string[] searchPaths = new string[]
                    {
                        System.IO.Path.GetDirectoryName(
                            doc.PathName ?? ""),
                        Environment.GetFolderPath(
                            Environment.SpecialFolder.Desktop),
                        Environment.GetFolderPath(
                            Environment.SpecialFolder.MyDocuments)
                    };

                    foreach (string dir in searchPaths)
                    {
                        if (string.IsNullOrEmpty(dir)) continue;
                        string full = System.IO.Path.Combine(dir, name);
                        if (!full.EndsWith(".dwg"))
                            full += ".dwg";
                        if (System.IO.File.Exists(full))
                            return full;
                    }
                }
            }
            catch { }

            try
            {
                Parameter p = import.get_Parameter(
                    BuiltInParameter.IMPORT_SYMBOL_NAME);
                string fileName = p != null ? p.AsString() : "DWG file";

                var dialog = new Microsoft.Win32.OpenFileDialog();
                dialog.Title = "Locate: " + fileName;
                dialog.Filter = "DWG files (*.dwg)|*.dwg";
                if (dialog.ShowDialog() == true)
                    return dialog.FileName;
            }
            catch { }

            return null;
        }

        private class DwgTextData
        {
            public string Text;
            public XYZ Position;
            public double HeightMm;
            public double Rotation;

            public double RawX;
            public double RawY;
            public double RawHeight;
            public double RawRotation;

            public DwgTextData(string text, XYZ pos, double hMm, double rot)
            {
                Text = text;
                Position = pos;
                HeightMm = hMm;
                Rotation = rot;
            }
        }

        // ========== LINE HELPERS ==========

        private void GetCurves(Document doc, GeometryElement geoElem,
            List<(Curve curve, GraphicsStyle style)> curveData)
        {
            foreach (GeometryObject geoObj in geoElem)
            {
                GraphicsStyle gs =
                    doc.GetElement(geoObj.GraphicsStyleId) as GraphicsStyle;

                if (geoObj is Curve curve)
                {
                    curveData.Add((curve, gs));
                }
                else if (geoObj is PolyLine polyline)
                {
                    IList<XYZ> pts = polyline.GetCoordinates();
                    for (int i = 0; i < pts.Count - 1; i++)
                    {
                        if (pts[i].DistanceTo(pts[i + 1]) < 0.003) continue;
                        curveData.Add((
                            Autodesk.Revit.DB.Line.CreateBound(
                                pts[i], pts[i + 1]), gs));
                    }
                }
                else if (geoObj is GeometryInstance geoInst)
                {
                    GetCurves(doc, geoInst.GetInstanceGeometry(), curveData);
                }
                else if (geoObj is GeometryElement nestedElem)
                {
                    GetCurves(doc, nestedElem, curveData);
                }
            }
        }

        private string GetHmvLineStyleName(Document doc, GraphicsStyle style)
        {
            if (style == null) return null;
            Category cat = style.GraphicsStyleCategory;
            if (cat == null) return null;

            int? weight = cat.GetLineWeight(GraphicsStyleType.Projection);
            int w = weight.HasValue ? weight.Value : 1;

            ElementId patternId =
                cat.GetLinePatternId(GraphicsStyleType.Projection);
            string patternName = "SOLID";
            if (patternId != null
                && patternId != ElementId.InvalidElementId
                && patternId.IntegerValue > 0)
            {
                LinePatternElement pat =
                    doc.GetElement(patternId) as LinePatternElement;
                if (pat != null) patternName = pat.Name.ToUpper();
            }
            return "HMV_LINEA " + patternName + " " + w;
        }

        private GraphicsStyle CreateLineStyle(Document doc,
            Category linesCat, GraphicsStyle sourceStyle, string name)
        {
            try
            {
                Category newCat = doc.Settings.Categories
                    .NewSubcategory(linesCat, name);
                Category srcCat = sourceStyle.GraphicsStyleCategory;

                int? weight = srcCat.GetLineWeight(
                    GraphicsStyleType.Projection);
                if (weight.HasValue)
                    newCat.SetLineWeight(weight.Value,
                        GraphicsStyleType.Projection);

                ElementId patternId = srcCat.GetLinePatternId(
                    GraphicsStyleType.Projection);
                if (patternId != null
                    && patternId != ElementId.InvalidElementId)
                    newCat.SetLinePatternId(patternId,
                        GraphicsStyleType.Projection);

                newCat.LineColor = new Autodesk.Revit.DB.Color(0, 0, 0);
                return newCat.GetGraphicsStyle(GraphicsStyleType.Projection);
            }
            catch { return null; }
        }

        private double SnapToAllowed(double sizeMm)
        {
            double closest = AllowedSizes[0];
            double minDiff = Math.Abs(sizeMm - closest);
            foreach (double a in AllowedSizes)
            {
                double diff = Math.Abs(sizeMm - a);
                if (diff < minDiff) { minDiff = diff; closest = a; }
            }
            return closest;
        }

        private TextNoteType CreateTextType(Document doc,
            TextNoteType baseType, string name, double sizeMm, string font)
        {
            try
            {
                TextNoteType nt = baseType.Duplicate(name) as TextNoteType;
                double sizeFeet = sizeMm / 304.8;
                nt.get_Parameter(BuiltInParameter.TEXT_SIZE).Set(sizeFeet);
                nt.get_Parameter(BuiltInParameter.TEXT_FONT).Set(font);
                nt.get_Parameter(BuiltInParameter.TEXT_WIDTH_SCALE).Set(1.0);
                nt.get_Parameter(BuiltInParameter.TEXT_STYLE_BOLD).Set(0);
                nt.get_Parameter(BuiltInParameter.TEXT_STYLE_ITALIC).Set(0);
                nt.get_Parameter(BuiltInParameter.TEXT_BACKGROUND).Set(1);
                return nt;
            }
            catch { return null; }
        }

        private bool IsStandardRevitLineStyle(string name)
        {
            string[] standard = new string[]
            {
                "Thin Lines", "Medium Lines", "Wide Lines", "Lines",
                "<Thin Lines>", "<Medium Lines>", "<Wide Lines>",
                "<Overhead>", "<Beyond>", "<Hidden>", "<Centerline>",
                "<Demolished>", "Axis of Rotation", "Hidden Lines",
                "Demolished", "Overhead", "Beyond", "Centerline",
                "Lineas finas", "Lineas medias", "Lineas anchas",
                "Lineas ocultas", "Eje de rotacion", "Demolido", "Encima"
            };

            foreach (string s in standard)
            {
                if (name.Equals(s, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }
    }

    public enum DwgConvertAction
    {
        ConvertLines,
        StandardizeTexts
    }
}