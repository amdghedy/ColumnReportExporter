using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace ColumnReport
{
    [Transaction(TransactionMode.Manual)]
    public class ColumnReport : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            using (MainForm form = new MainForm())
            {
                if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string filePath = form.FilePath;
                    if (string.IsNullOrEmpty(filePath))
                    {
                        message = "File path not provided.";
                        return Result.Failed;
                    }

                    // Initialize exporter
                    ColumnReportExporter exporter = new ColumnReportExporter();
                    // Execute export
                    exporter.Execute(commandData, filePath);

                    return Result.Succeeded;
                }
                else
                {
                    message = "User cancelled the operation.";
                    return Result.Cancelled;
                }
            }
        }
    }
}
