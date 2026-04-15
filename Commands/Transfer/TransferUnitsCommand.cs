using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HMVTools
{
    [Transaction(TransactionMode.Manual)]
    public class TransferUnitsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document srcDoc = uiApp.ActiveUIDocument.Document;

            // Gather currently open documents (excluding the active source and families)
            List<TargetDocEntry> openDocs = new List<TargetDocEntry>();
            foreach (Document doc in uiApp.Application.Documents)
            {
                if (!doc.IsLinked && !doc.IsFamilyDocument && doc.PathName != srcDoc.PathName)
                {
                    openDocs.Add(new TargetDocEntry
                    {
                        Title = doc.Title,
                        PathName = doc.PathName,
                        OpenDoc = doc,
                        IsOpenInRevit = true
                    });
                }
            }

            var win = new TransferUnitsWindow(srcDoc.Title, openDocs);

            if (win.ShowDialog() != true || win.GetSelectedTargets().Count == 0)
                return Result.Cancelled;

            List<TargetDocEntry> targetsToProcess = win.GetSelectedTargets();

           
            // 2. Execute Batch Process
            TransferUnitsResult result = TransferUnitsManager.ProcessBatch(uiApp.Application, srcDoc, targetsToProcess);

            // 3. Show Report
            TaskDialog.Show("HMV Tools - Units Transfer Report", result.BuildReport());

            return Result.Succeeded;
        }
    }
}