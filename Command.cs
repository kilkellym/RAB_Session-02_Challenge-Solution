#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

#endregion

namespace RAB_Session_02_Challenge_Solution
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        internal double ConvertMetersToFeet(double meters)
        {
            double feet = meters * 3.28084;

            return feet;
        }
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // 1. declare variables
            string levelPath = "C:\\Users\\micha\\OneDrive\\Documents\\Revit Add-in Bootcamp\\RAB_Session_02_Challenge_Levels.csv";
            string sheetPath = "C:\\Users\\micha\\OneDrive\\Documents\\Revit Add-in Bootcamp\\RAB_Session_02_Challenge_Sheets.csv";

            string testPath = Path.GetDirectoryName(doc.PathName);
            string testPath2 = testPath + "\\Excel File\\";

            List<string[]> levelData = new List<string[]>();
            List<string[]> sheetData = new List<string[]>();

            // 2. read text files
            string[] levelArray = File.ReadAllLines(levelPath);
            string[] sheetArray = File.ReadAllLines(sheetPath);

            // 3. loop through file data and put into list
            foreach(string levelString in levelArray)
            {
                string[] cellData = levelString.Split(',');

                levelData.Add(cellData);
            }

            foreach (string sheetString in sheetArray)
            {
                string[] cellData = sheetString.Split(',');

                sheetData.Add(cellData);
            }

            // 4. remove header rows
            levelData.RemoveAt(0);
            sheetData.RemoveAt(0);

            // 5. create levels
            Transaction t1 = new Transaction(doc);
            t1.Start("Create levels");

            foreach (string[] currentLevelData in levelData)
            {
                string stringHeight = currentLevelData[1];
                double heightFeet = 0;
                bool convertFeet = double.TryParse(stringHeight, out heightFeet);

                if(convertFeet == false)
                {
                    TaskDialog.Show("Error", "Could not convert value. Defaulting to 0");
                }

                string stringMeters = currentLevelData[2];
                double heightMeter = 0;
                bool convertMeters = double.TryParse(stringMeters, out heightMeter);
                double metersToFeet = ConvertMetersToFeet(heightMeter);

                Level currentLevel = Level.Create(doc, metersToFeet);
                currentLevel.Name = currentLevelData[0];
            }

            t1.Commit();
            t1.Dispose();

            // 6. get title block element ID
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
            ElementId tblockId = collector.FirstElementId();

            // 7. create sheets
            Transaction t2 = new Transaction(doc);
            t2.Start("Create sheets");

            foreach (string[] currentSheetData in sheetData)
            {
                ViewSheet currentSheet = ViewSheet.Create(doc, tblockId);
                currentSheet.SheetNumber = currentSheetData[0];
                currentSheet.Name = currentSheetData[1];
            }

            t2.Commit();
            t2.Dispose();

            return Result.Succeeded;
        }


    }
}
