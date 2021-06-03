#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ExcelDataReader;
using System;
using System.IO;

#endregion

namespace RevitLevel
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        private const string LIST = @"C:\Users\џрырыр\source\repos\RevitLevel\Create levels from Excel.xlsx";

        [Obsolete]
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            FilteredElementCollector newFilter = new FilteredElementCollector(doc);

            ExcelReader(doc);
            return Result.Succeeded;
        }

        private void ExcelReader(Document doc)
        {
            using (var stream = File.Open(LIST, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        int i = 0;
                        while (reader.Read())
                        {
                            if (i != 0)
                            {
                                var nameFloor = reader.GetValue(0);
                                var elevation = reader.GetDouble(1);
                                CreateLevel(doc, nameFloor.ToString(), elevation);
                            }
                            i++;
                        }
                    } while (reader.NextResult());

                    var result = reader.AsDataSet();
                }
            }
        }

        private Result CreateLevel(Autodesk.Revit.DB.Document document, string name, double elevation)
        {
            using (Transaction firstTrans = new Transaction(document))
            {
                try
                {
                    firstTrans.Start("Start");
                    Level level = Level.Create(document, elevation);

                    if (null == level)
                    {
                        throw new Exception("Create a new level failed.");
                    }
                    level.Name = name;

                    firstTrans.Commit();
                    return Result.Succeeded;
                }
                catch
                {
                    return Result.Failed;
                }
            }

        }
    }
}
