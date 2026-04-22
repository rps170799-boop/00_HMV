using System.Collections.Generic;
using Autodesk.Revit.DB;

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

    /// <summary>Represents a sheet available for transfer.</summary>
    public class SheetEntry
    {
        public int Id { get; set; }
        public string SheetNumber { get; set; }
        public string SheetName { get; set; }
        public int ViewportCount { get; set; }
        public bool IsPlaceholder { get; set; }
    }

    /// <summary>Controls how views are handled when a match exists in the target.</summary>
    public enum ViewTransferMode
    {
        /// <summary>
        /// If a view with the same name exists, clear its contents and
        /// re-copy everything from source (pyRevit-stable approach).
        /// </summary>
        Update,

        /// <summary>
        /// Always create a new view in the target, even if one with the
        /// same name already exists (appends " (2)", " (3)", etc.).
        /// </summary>
        Create
    }

    /// <summary>Options controlling what gets copied with sheets.</summary>
    public class SheetCopyOptions
    {
        public bool CopyViewports { get; set; } = true;
        public bool CopySchedules { get; set; } = true;
        public bool CopyTitleblock { get; set; } = true;
        public bool CopyRevisions { get; set; } = false;
        public bool CopyGuideGrids { get; set; } = true;
        public bool PlaceholdersAsSheets { get; set; } = true;
        public bool PreserveDetailNumbers { get; set; } = true;
    }

    /// <summary>User selections returned from the migration window.</summary>
    public class MigrationSettings
    {
        public int TargetDocIndex { get; set; }
        public List<int> SelectedViewIds { get; set; } = new List<int>();
        public List<int> SelectedSheetIds { get; set; } = new List<int>();
        public bool IncludeAnnotations { get; set; } = true;
        public bool IncludeRefMarkers { get; set; } = true;

        /// <summary>Create new views vs update existing ones.</summary>
        public ViewTransferMode TransferMode { get; set; }
            = ViewTransferMode.Create;

        /// <summary>Sheet-specific copy options.</summary>
        public SheetCopyOptions SheetOptions { get; set; }
            = new SheetCopyOptions();
    }

    /// <summary>Represents a target document for the Transfer Units tool.</summary>
    public class TargetDocEntry
    {
        public string Title { get; set; }
        public string PathName { get; set; }
        public Document OpenDoc { get; set; }
        public bool IsOpenInRevit { get; set; }
        public bool IsSelected { get; set; } = true;
    }
}