using System.Collections.Generic;

namespace HMVTools
{
    /// <summary>Represents an open Revit document for the target dropdown.</summary>
    public class OpenDocEntry
    {
        public string Title { get; set; }
        public string PathName { get; set; }
        public int Index { get; set; }
    }

    /// <summary>Represents a view available for transfer.</summary>
    public class ViewEntry
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Category { get; set; }
    }

    /// <summary>User selections returned from the migration window.</summary>
    public class MigrationSettings
    {
        public int TargetDocIndex { get; set; }
        public List<int> SelectedViewIds { get; set; } = new List<int>();
        public bool IncludeAnnotations { get; set; } = true;
        public bool IncludeRefMarkers { get; set; } = true;
    }
}