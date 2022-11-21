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
using Forms = System.Windows.Forms;

#endregion

namespace RAB_Session_02_Challenge_Solution
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
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
            string levelPath = "";
            string sheetPath = "";

            Forms.OpenFileDialog levelFile = new Forms.OpenFileDialog();
            levelFile.InitialDirectory = "C:\\";
            levelFile.Multiselect = false;
            levelFile.Filter = "CSV Files|*.csv";

            Forms.OpenFileDialog sheetFile = new Forms.OpenFileDialog();
            sheetFile.InitialDirectory = "C:\\";
            sheetFile.Multiselect = false;
            sheetFile.Filter = "CSV Files|*.csv";

            if(levelFile.ShowDialog() == Forms.DialogResult.OK)
            {
                levelPath = levelFile.FileName;
            }

            if(sheetFile.ShowDialog() == Forms.DialogResult.OK)
            {
                sheetPath = sheetFile.FileName;
            }

            // 2. read text files
            List<LevelData> levelDataList = new List<LevelData>();
            List<SheetData> sheetDataList = new List<SheetData>();

            string[] levelArray = File.ReadAllLines(levelPath);
            string[] sheetArray = File.ReadAllLines(sheetPath);

            // 3. loop through file data and put into list
            foreach(string levelString in levelArray)
            {
                string[] cellData = levelString.Split(',');

                LevelData currentLevelData = new LevelData();
                currentLevelData.LevelName = cellData[0];
                currentLevelData.HeightFeet = ConvertStringToDouble(cellData[1]);
                currentLevelData.HeightMeters = ConvertStringToDouble(cellData[2]);

                levelDataList.Add(currentLevelData);
            }

            foreach (string sheetString in sheetArray)
            {
                string[] cellData = sheetString.Split(',');

                SheetData currentSheetData = new SheetData();
                currentSheetData.SheetNumber = cellData[0];
                currentSheetData.SheetName = cellData[1];

                sheetDataList.Add(currentSheetData);
            }

            // 4. remove header rows
            levelDataList.RemoveAt(0);
            sheetDataList.RemoveAt(0);

            // 5. create levels
            Transaction t1 = new Transaction(doc);
            t1.Start("Create levels");

            foreach (LevelData currentLevelData in levelDataList)
            {
                // OPTIONAL - if you're using the metric system
                //double metersToFeet = ConvertMetersToFeet(currentLevelData.HeightMeters);

                Level currentLevel = Level.Create(doc, currentLevelData.HeightFeet);
                currentLevel.Name = currentLevelData.LevelName;

                ViewFamilyType planVFT = GetViewFamilyTypeByName(doc, "Floor Plan", ViewFamily.FloorPlan);
                ViewFamilyType ceilingPlanVFT = GetViewFamilyTypeByName(doc, "Ceiling Plan", ViewFamily.CeilingPlan);

                // create floor plan and ceiling plan views
                ViewPlan currentFloorPlan = ViewPlan.Create(doc, planVFT.Id, currentLevel.Id);
                ViewPlan currentCeilingPlan = ViewPlan.Create(doc, ceilingPlanVFT.Id, currentLevel.Id);
                
                if(currentFloorPlan.Name.Contains("Roof") == true)
                {
                    currentFloorPlan.Name = "Roof Plan";
                }    
                else
                {
                    currentFloorPlan.Name = currentFloorPlan.Name + " Floor Plan";
                }
                
                currentCeilingPlan.Name = currentCeilingPlan.Name + " RCP";
            }

            t1.Commit();
            t1.Dispose();

            // 6. get title block element ID
            Element tblock = GetTitleBlockByName(doc, "E1 30x42 Horizontal");

            // 7. create sheets
            Transaction t2 = new Transaction(doc);
            t2.Start("Create sheets");

            foreach (SheetData currentSheetData in sheetDataList)
            {
                ViewSheet currentSheet = ViewSheet.Create(doc, tblock.Id);

                currentSheet.SheetNumber = currentSheetData.SheetNumber;
                currentSheet.Name = currentSheetData.SheetName;

                View currentView = GetViewByName(doc, currentSheet.Name);

                XYZ insPoint = new XYZ(1.5, 1, 0);
                //XYZ insPoint = GetSheetCenterPoint(currentSheet);

                Viewport currentVP = Viewport.Create(doc, currentSheet.Id, currentView.Id, insPoint);
            }

            t2.Commit();
            t2.Dispose();

            return Result.Succeeded;
        }

        private XYZ GetSheetCenterPoint(ViewSheet currentSheet)
        {
            // Get the middle point of the sheet (insertion point)
            BoundingBoxUV outline = currentSheet.Outline;
            double x = (outline.Max.U + outline.Min.U) / 2;
            double y = (outline.Max.V + outline.Min.V) / 2;

            XYZ returnPoint = new XYZ(x, y, 0);

            return returnPoint;
        }

        private View GetViewByName(Document doc, string name)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Views);

            foreach(View currentView in collector)
            {
                if(currentView.Name == name)
                {
                    return currentView;
                }    
            }

            return null;
        }

        private ViewFamilyType GetViewFamilyTypeByName(Document doc, string vftName, ViewFamily vf)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewFamilyType));

            foreach(ViewFamilyType currentVFT in collector)
            {
                if(currentVFT.Name == vftName && currentVFT.ViewFamily == vf)
                {
                    return currentVFT;
                }
            }

            return null;
        }

        internal Element GetTitleBlockByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);

            foreach(Element currentTblock in collector)
            {
                if(currentTblock.Name == typeName)
                {
                    return currentTblock;
                }
            }

            return null;
        }

        private double ConvertStringToDouble(string numberString)
        {
            double height = 0;
            bool convertString = double.TryParse(numberString, out height);

            return height;
        }
        internal double ConvertMetersToFeet(double meters)
        {
            double feet = meters * 3.28084;

            return feet;
        }

        public struct LevelData
        {
            public string LevelName;
            public double HeightFeet;
            public double HeightMeters;
        }

        public struct SheetData
        {
            public string SheetNumber;
            public string SheetName;
        }


    }
}
