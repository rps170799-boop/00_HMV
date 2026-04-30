using System.Collections.Generic;

namespace HMVTools
{
    public class TransferResult
    {
        public int ElementsCopied { get; set; }
        public int ViewsCreated { get; set; }
        public int ViewsUpdated { get; set; }
        public int SheetsCreated { get; set; }
        public int SheetsUpdated { get; set; }
        public int ViewportsCopied { get; set; }
        public int TemplatesRemapped { get; set; }
        public int RefMarkersNoted { get; set; }
        public int AnnotationsCopied { get; set; }
        public List<string> Warnings { get; set; } = new List<string>();
        public List<string> Errors { get; set; } = new List<string>();

        public string BuildReport()
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("═══  MIGRATION REPORT  ═══");
            sb.AppendLine();
            sb.AppendLine($"  Model elements copied:    {ElementsCopied}");
            sb.AppendLine($"  Views created:            {ViewsCreated}");
            sb.AppendLine($"  Views updated:            {ViewsUpdated}");
            sb.AppendLine($"  Sheets created:           {SheetsCreated}");
            sb.AppendLine($"  Sheets updated:           {SheetsUpdated}");
            sb.AppendLine($"  Viewports placed:         {ViewportsCopied}");
            sb.AppendLine($"  Templates remapped:       {TemplatesRemapped}");
            sb.AppendLine($"  Ref markers identified:   {RefMarkersNoted}");
            sb.AppendLine($"  Annotations copied:       {AnnotationsCopied}");

            if (Warnings.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("── WARNINGS ──");
                foreach (var w in Warnings)
                    sb.AppendLine($"  ⚠  {w}");
            }
            if (Errors.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("── ERRORS ──");
                foreach (var e in Errors)
                    sb.AppendLine($"  ✗  {e}");
            }
            return sb.ToString();
        }
    }
}
