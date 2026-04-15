using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.DB;

namespace HMVTools
{
    // Preprocessor stays the same
    public class SuppressWarningsPreprocessor : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            IList<FailureMessageAccessor> failures = failuresAccessor.GetFailureMessages();
            foreach (FailureMessageAccessor failure in failures)
            {
                if (failure.GetSeverity() == FailureSeverity.Warning)
                    failuresAccessor.DeleteWarning(failure);
            }
            return FailureProcessingResult.Continue;
        }
    }

    public static class TransferUnitsManager
    {
        public static TransferUnitsResult ProcessBatch(Autodesk.Revit.ApplicationServices.Application app, Document srcDoc, List<TargetDocEntry> targets)
        {
            var result = new TransferUnitsResult();
            Units srcUnits = srcDoc.GetUnits();
            IList<ForgeTypeId> modifiableSpecs = Units.GetModifiableSpecs();

            result.TotalSpecsFound = modifiableSpecs.Count;

            foreach (TargetDocEntry target in targets)
            {
                Document tgtDoc = target.OpenDoc;
                bool isHeadless = false;

                // Reset counters for this specific file run
                int successfulSpecs = 0;
                int failedSpecs = 0;

                try
                {
                    if (tgtDoc == null)
                    {
                        if (!File.Exists(target.PathName))
                        {
                            result.Errors.Add($"File not found: {target.Title}");
                            continue;
                        }

                        ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(target.PathName);
                        OpenOptions openOpts = new OpenOptions();

                        BasicFileInfo fileInfo = BasicFileInfo.Extract(target.PathName);
                        if (fileInfo.IsWorkshared)
                        {
                            openOpts.DetachFromCentralOption = DetachFromCentralOption.DetachAndDiscardWorksets;
                        }

                        tgtDoc = app.OpenDocumentFile(modelPath, openOpts);
                        isHeadless = true;
                    }

                    using (Transaction t = new Transaction(tgtDoc, "HMV - Transfer Project Units"))
                    {
                        FailureHandlingOptions fho = t.GetFailureHandlingOptions();
                        fho.SetFailuresPreprocessor(new SuppressWarningsPreprocessor());
                        fho.SetClearAfterRollback(true);
                        t.SetFailureHandlingOptions(fho);

                        t.Start();

                        Units tgtUnits = tgtDoc.GetUnits();

                        // 1. Transfer Global Settings
                        tgtUnits.DecimalSymbol = srcUnits.DecimalSymbol;
                        tgtUnits.DigitGroupingSymbol = srcUnits.DigitGroupingSymbol;
                        tgtUnits.DigitGroupingAmount = srcUnits.DigitGroupingAmount;

                        // 2. Transfer Individual Unit Specifications Safely
                        foreach (ForgeTypeId specTypeId in modifiableSpecs)
                        {
                            try
                            {
                                // SIMPLIFIED: Directly grab and assign the FormatOptions object.
                                // This avoids crashing on units that have no symbols or weird edge cases.
                                FormatOptions srcFormat = srcUnits.GetFormatOptions(specTypeId);
                                tgtUnits.SetFormatOptions(specTypeId, srcFormat);

                                successfulSpecs++;
                            }
                            catch (Exception specEx)
                            {
                                failedSpecs++;
                                // Only log the warning once so we don't spam the report 60 times
                                string warningMsg = $"Could not copy spec '{specTypeId.TypeId}': {specEx.Message}";
                                if (!result.SpecWarnings.Contains(warningMsg))
                                {
                                    result.SpecWarnings.Add(warningMsg);
                                }
                            }
                        }

                        tgtDoc.SetUnits(tgtUnits);
                        t.Commit();
                    }

                    if (isHeadless)
                    {
                        SaveAsOptions saveOpts = new SaveAsOptions { OverwriteExistingFile = true };
                        tgtDoc.SaveAs(target.PathName, saveOpts);
                    }

                    result.FilesSuccessfullyProcessed++;

                    // Update global tallies based on this file's attempt
                    result.SpecsSuccessfullyMigratedPerFile = successfulSpecs;
                    result.SpecsFailedPerFile = failedSpecs;
                }
                catch (Exception ex)
                {
                    result.Errors.Add($"Failed processing '{target.Title}': {ex.Message}");
                }
                finally
                {
                    if (isHeadless && tgtDoc != null && tgtDoc.IsValidObject)
                    {
                        tgtDoc.Close(false);
                    }
                }
            }

            return result;
        }
    }
}