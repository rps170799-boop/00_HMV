# TransferManager Refactoring Summary

Original file: `Helpers\Transfermanager.cs`  
All logic, method signatures, and behavior are **unchanged** — only the physical location of each method moved.  
The project must compile successfully after every step.

---

## Before → After: File Structure

| Before | After |
|--------|-------|
| Single `TransferManager.cs` (~3 400 lines) | 9 focused files, each < 500 lines |

---

## Step 2a — `UseDestinationTypesHandler.cs`

Extracted class implementing `IDuplicateTypeNamesHandler`.  
Tells Revit to reuse types that already exist in the target document instead of creating duplicates.

| Method | Access |
|--------|--------|
| `Execute(...)` | `public` |

---

## Step 2b — `SwallowErrorsHandler.cs`

Extracted class implementing `IFailuresPreprocessor`.  
Swallows non-fatal Revit warnings during transactions, equivalent to pyRevit `swallow_errors=True`.

| Method | Access |
|--------|--------|
| `PreprocessFailures(...)` | `public` |

---

## Step 2c — `TransferResult.cs`

Extracted result/counter class used to accumulate what happened during a migration.

| Member | Type |
|--------|------|
| `ElementsCopied` | `int` |
| `ViewsCreated` | `int` |
| `ViewsUpdated` | `int` |
| `SheetsCreated` | `int` |
| `SheetsUpdated` | `int` |
| `TemplatesRemapped` | `int` |
| `AnnotationsCopied` | `int` |
| `Warnings` | `List<string>` |
| `Errors` | `List<string>` |
| `BuildReport()` | `public string` |

---

## Step 2d — `ViewportData.cs`

Extracted DTO (data transfer object) capturing viewport placement data before and after copy.

| Member | Type | Notes |
|--------|------|-------|
| `BoxCenter` | `XYZ` | Centre of the viewport box |
| `LabelOffset` | `XYZ` | Offset of the view title label |
| `LabelLineLength` | `double?` | Length of the label leader line (`null` if not set) |
| `BBoxMin` | `XYZ` | Minimum corner of the viewport bounding box |
| `BBoxMax` | `XYZ` | Maximum corner of the viewport bounding box |
| `Rotation` | `ViewportRotation` | Viewport rotation (defaults to `None`) |

---

## Step 3 — `ViewContentCopier.cs`

Extracted view-content transfer helpers. Called by `ViewFactory` and `TransferManager`.

| Method | Access | Purpose |
|--------|--------|---------|
| `GetViewContents(source, view, opts)` | `public` | Collects view-owned element IDs (annotations, dims, detail items) |
| `ClearViewContents(target, view, opts)` | `public` | Deletes existing view-owned elements before an update |
| `CopyViewContentsViewToView(source, srcView, target, tgtView, opts)` | `public` | View-to-view copy of owned elements |
| `CopyViewProperties(srcView, tgtView)` | `public` | Copies scale, detail level, display style, description |

---

## Step 4 — `ViewportPlacer.cs`

Extracted viewport placement logic. Called from `TransferManager.CopySingleSheet`.

| Method | Access | Purpose |
|--------|--------|---------|
| `CopySheetViewports(source, srcSheet, target, destSheet, transform, viewMap, result, settings)` | `public` | Top-level: iterates viewports on a sheet and places them on the destination sheet |
| `ApplyViewportType(target, viewport, typeName)` | `public` | Matches and sets the viewport type by name |
| `ApplyDetailNumber(viewport, detailNumber)` | `public` | Sets the detail number parameter on a placed viewport |
| `CaptureViewportData(source, viewport)` | `public` | Snapshots position/label/type data before copy |
| `ApplyLabelProperties(target, viewport, data)` | `public` | Re-applies label offset after placement |
| `CorrectViewportByBBox(target, viewport, data)` | `public` | Nudges viewport position using bounding-box delta |
| `GetViewSheetNumber(view)` | `public` | Reads the sheet-number parameter from a view |

---

## Step 5 — `ViewFactory.cs`

Extracted view creation, duplication, batch copy, and matching logic.

| Method | Access | Purpose |
|--------|--------|---------|
| `FindMatchingView(target, sourceView)` | `public` | Finds a view in target by ViewType + Name (+ SheetNumber for sheets) |
| `CopyViewBatch(source, target, ids, transform, opts, viewMap, result, batchLabel, mode)` | `internal` | PASS 1/2 batch copy: regular views inside a transaction, ViewSection views outside |
| `CopyOrUpdateSingleView(source, target, srcView, transform, result, settings)` | `internal` | Dispatches by ViewType: update existing or create new view |
| `ForceCreateView(source, target, srcViewId, result)` | `private` | Dispatcher to ForceCreatePlan / ForceCreateSection |
| `ForceCreatePlan(source, target, srcPlan, result)` | `private` | Duplicates an existing plan on the nearest level, copies crop/range/properties |
| `ForceCreateSection(source, target, srcSection, result)` | `private` | Duplicates a donor section, moves and rotates the viewer element to match source position |
| `FindClosestLevel(doc, elevation)` | `private` | Finds the level in a document closest to a given elevation |
| `GetUniqueViewName(doc, desired)` | `internal` | Returns a name that does not conflict with existing view names |
| `SetSwallowErrors(transaction)` | `private` | Attaches `SwallowErrorsHandler` to a transaction |

> **Note:** `FindClosestLevel` and `SetSwallowErrors` are generic document/transaction utilities that happen to live in `ViewFactory` because they are only used by `ForceCreatePlan` and `ForceCreateSection`. They are `private` so there is no leakage. A future cleanup could move them to a shared `TransactionUtils` or `DocumentQueryUtils` helper, but this is not blocking.

---

## Step 6 — `SheetHelpers.cs`

Extracted shared-coordinate utilities, guide grid copy, and revision copy.  
External callers in `MigrateElementsCommand.cs` updated from `TransferManager.*` to `SheetHelpers.*`.

| Method | Access | Purpose |
|--------|--------|---------|
| `ComputeSharedCoordinateTransform(source, target)` | `public` | Computes the transform between two documents' shared coordinate systems |
| `ValidateTransform(transform, out warning)` | `public` | Checks for NaN and warns if translation exceeds 10 km |
| `CopySheetGuides(source, srcSheet, target, destSheet, result)` | `internal` | Copies and assigns guide grids between sheets |
| `FindGuideByName(doc, guideName)` | `private` | Finds a guide grid element by name |
| `CopySheetRevisions(source, srcSheet, target, destSheet, result)` | `internal` | Copies revision entries from source sheet to destination sheet |
| `FindOrCreateRevision(srcRev, allDestRevs, target, result)` | `private` | Finds a matching revision in target or creates a new one |

---

## Steps 7+8 — `ViewTransferHelper.cs`

Extracted view template assignment, annotation copy, graphic override copy, and reference marker identification.  
External callers in `MigrateElementsCommand.cs` updated from `TransferManager.*` to `ViewTransferHelper.*`.

| Method | Access | Purpose |
|--------|--------|---------|
| `CopyAndAssignViewTemplates(source, target, viewMap, result)` | `public` | Copies view templates from source and assigns them to mapped target views |
| `CopyViewAnnotations(source, target, viewMap, result)` | `public` | 2nd-pass copy of annotations, dimensions, text, detail items for all mapped views |
| `TryCopyAnnotations(srcView, tgtView, ids, opts, result, label)` | `private` | Single-category view-to-view annotation copy with rollback on failure |
| `CopyCategoryOverrides(source, target, viewMap, result)` | `public` | Copies per-category graphic overrides for all mapped views |
| `RemapOverride(source, target, src, fillMap, lineMap, copyOpts)` | `private` | Rebuilds an `OverrideGraphicSettings` object with remapped pattern IDs |
| `RemapFillPattern(source, target, srcPatternId, cache, copyOpts)` | `private` | Finds or copies a fill pattern to the target document |
| `RemapLinePattern(source, target, srcPatternId, cache, copyOpts)` | `private` | Finds or copies a line pattern to the target document |
| `IsOverrideEmpty(ogs)` | `private` | Returns `true` if all override fields are default (nothing to apply) |
| `ColorsEqual(a, b)` | `private` | Null-safe RGB color comparison |
| `IdentifyReferenceMarkers(source, sourceView)` | `public` | Lists section/callout/elevation reference markers visible in a view |
| `CollectRefs(doc, view, label, markers, getter)` | `private` | Helper that invokes a view reference getter and formats results |

---

## What Stayed in `TransferManager.cs`

Only the top-level public entry points that external commands call directly.

| Method | Purpose |
|--------|---------|
| `CopyModelElements(source, target, elementIds, transform, result, groupName)` | Copies 3D model elements and groups them in the target |
| `CopyViews(source, target, viewIds, transform, result, mode)` | Orchestrates view copy/update by type (plan, section, drafting, legend) |
| `CopySheets(source, target, sheetIds, transform, result, settings)` | Iterates sheets and delegates to `CopySingleSheet` |
| `CopySingleSheet(source, target, srcSheet, transform, viewMap, result, settings)` | Creates or updates one sheet, then places viewports, annotations, revisions |

---

## Call-Site Updates Summary

| Caller file | Old call | New call |
|-------------|----------|----------|
| `MigrateElementsCommand.cs` | `TransferManager.ComputeSharedCoordinateTransform` | `SheetHelpers.ComputeSharedCoordinateTransform` |
| `MigrateElementsCommand.cs` | `TransferManager.ValidateTransform` | `SheetHelpers.ValidateTransform` |
| `MigrateElementsCommand.cs` | `TransferManager.CopyAndAssignViewTemplates` | `ViewTransferHelper.CopyAndAssignViewTemplates` |
| `MigrateElementsCommand.cs` | `TransferManager.CopyViewAnnotations` | `ViewTransferHelper.CopyViewAnnotations` |
| `MigrateElementsCommand.cs` | `TransferManager.CopyCategoryOverrides` | `ViewTransferHelper.CopyCategoryOverrides` |
| `MigrateElementsCommand.cs` | `TransferManager.IdentifyReferenceMarkers` | `ViewTransferHelper.IdentifyReferenceMarkers` |
| `TransferManager.cs` | `CopyViewBatch(...)` | `ViewFactory.CopyViewBatch(...)` |
| `TransferManager.cs` | `CopyOrUpdateSingleView(...)` | `ViewFactory.CopyOrUpdateSingleView(...)` |
| `TransferManager.cs` | `FindMatchingView(...)` | `ViewFactory.FindMatchingView(...)` |
| `TransferManager.cs` | `CopySheetRevisions(...)` | `SheetHelpers.CopySheetRevisions(...)` |
| `TransferManager.cs` | `CopyViewAnnotations(...)` | `ViewTransferHelper.CopyViewAnnotations(...)` |
| `ViewportPlacer.cs` | `TransferManager.CopyOrUpdateSingleView(...)` | `ViewFactory.CopyOrUpdateSingleView(...)` |
