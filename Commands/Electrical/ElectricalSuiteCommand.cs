using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HMVTools
{
    // ═══════════════════════════════════════════════════════════════════
    //  ELECTRICAL SUITE  —  single ribbon entry point, opens the launcher
    // ═══════════════════════════════════════════════════════════════════

    [Transaction(TransactionMode.Manual)]
    public class ElectricalSuiteCommand : IExternalCommand
    {
        private static ElectricalSuiteWindow _window = null;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (_window != null && _window.IsLoaded)
            {
                _window.Focus();
                return Result.Succeeded;
            }

            _window = new ElectricalSuiteWindow(commandData.Application);

            var helper = new System.Windows.Interop.WindowInteropHelper(_window);
            helper.Owner = commandData.Application.MainWindowHandle;

            _window.Show();
            return Result.Succeeded;
        }

        public static void ClearWindow() { _window = null; }
    }
}
