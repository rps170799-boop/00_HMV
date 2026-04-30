using Autodesk.Revit.DB;

namespace HMVTools
{
    /// <summary>
    /// Holds source viewport positioning data for accurate
    /// placement in the destination sheet.
    /// </summary>
    public class ViewportData
    {
        public XYZ BoxCenter { get; set; }
        public XYZ LabelOffset { get; set; }
        public double? LabelLineLength { get; set; }
        public XYZ BBoxMin { get; set; }
        public XYZ BBoxMax { get; set; }
        public ViewportRotation Rotation { get; set; } = ViewportRotation.None;
    }
}
