using System.Collections.Generic;

namespace HMVTools
{
    public class TransferUnitsResult
    {
        public int FilesSuccessfullyProcessed { get; set; } = 0;
        public int TotalSpecsFound { get; set; } = 0;
        public int SpecsSuccessfullyMigratedPerFile { get; set; } = 0;
        public int SpecsFailedPerFile { get; set; } = 0;
        public List<string> Errors { get; set; } = new List<string>();
        public List<string> SpecWarnings { get; set; } = new List<string>();

        public string BuildReport()
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("═══  UNITS TRANSFER REPORT  ═══");
            sb.AppendLine();
            sb.AppendLine($"  Models Processed:             {FilesSuccessfullyProcessed}");
            sb.AppendLine($"  Global Settings Migrated:     3 (Decimal symbol, grouping, etc.)");
            sb.AppendLine($"  Unit Specs Migrated (Each):   {SpecsSuccessfullyMigratedPerFile} / {TotalSpecsFound}");

            if (SpecsFailedPerFile > 0)
            {
                sb.AppendLine($"  Unit Specs Skipped (Each):    {SpecsFailedPerFile} (Check warnings below)");
            }

            if (Errors.Count > 0 || SpecWarnings.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("═══  WARNINGS / ERRORS  ═══");
                foreach (string err in Errors)
                {
                    sb.AppendLine($"  • {err}");
                }
                foreach (string warn in SpecWarnings)
                {
                    sb.AppendLine($"  • [Unit Skipped] {warn}");
                }
            }
            else
            {
                sb.AppendLine();
                sb.AppendLine("  Status: Completed without any errors.");
            }

            return sb.ToString();
        }
    }
}