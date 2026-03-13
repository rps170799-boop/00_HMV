using Autodesk.Revit.DB;

namespace HMVTools
{
    /// <summary>
    /// Replaces the external FamilyLoaderHelper.dll.
    /// Always overwrites parameters when reloading a family into the project.
    /// </summary>
    public class TextAuditFamilyLoadOptions : IFamilyLoadOptions
    {
        public bool OnFamilyFound(bool familyInUse,
            out bool overwriteParameterValues)
        {
            overwriteParameterValues = true;
            return true;   // continue loading
        }

        public bool OnSharedFamilyFound(Family sharedFamily,
            bool familyInUse, out FamilySource source,
            out bool overwriteParameterValues)
        {
            source = FamilySource.Family;
            overwriteParameterValues = true;
            return true;
        }
    }
}