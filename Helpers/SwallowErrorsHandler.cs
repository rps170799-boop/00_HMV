using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;

namespace HMVTools
{
    public class SwallowErrorsHandler : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(
            FailuresAccessor failuresAccessor)
        {
            var failures = failuresAccessor.GetFailureMessages();

            foreach (var f in failures)
            {
                // Delete warnings (non-critical)
                if (f.GetSeverity() == FailureSeverity.Warning)
                {
                    failuresAccessor.DeleteWarning(f);
                }
                // Try to resolve errors
                else if (f.HasResolutions())
                {
                    failuresAccessor.ResolveFailure(f);
                }
                else
                {
                    failuresAccessor.DeleteWarning(f);
                }
            }

            return FailureProcessingResult.Continue;
        }
    }
}
