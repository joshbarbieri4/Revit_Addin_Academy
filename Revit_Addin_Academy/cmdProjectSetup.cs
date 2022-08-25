#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;

#endregion

namespace Revit_Addin_Academy
{
	[Transaction(TransactionMode.Manual)]
	public class cmdProjectSetup : IExternalCommand
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

			Forms.OpenFileDialog dialog = new Forms.OpenFileDialog();
			dialog.InitialDirectory = @"C:\";
			dialog.Multiselect = false;
			dialog.Filter = "Excel Files | *.xls; *.xlsx; *.xlsm | All files | *.*"; // What the user sees & specifiy what type of file the user can select

			// A check point, if the user doesn't do what we want then stop code here
			if (dialog.ShowDialog() != Forms.DialogResult.OK)
			{
				TaskDialog.Show("Error", "Please select an Excel file.");
				return Result.Failed;
			}

			string excelFile = dialog.FileName;
			int levelCounter = 0;
			int sheetCounter = 0;

			try
			{
				// open Excel
				Excel.Application excelApp = new Excel.Application();
				Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);

				// Tip 1 to creating the bones of a method
				// if I righ click on the method GetExcelWorksheetByName - select Quick Actions and Refactorings... to create the bones of a method
				Excel.Worksheet excelWs1 = GetExcelWorksheetByName(excelWb, "Levels");
				Excel.Worksheet excelWs2 = GetExcelWorksheetByName(excelWb, "Sheets");

				// A method to read the levels/sheets data
				List<LevelStruct> levelData = GetLevelDataFromExcel(excelWs1);
				List<SheetStruct> sheetData = GetSheetDataFromExcel(excelWs2);

				excelWb.Close();
				excelApp.Quit();


				using (Transaction t = new Transaction(doc))
				{
					ViewFamilyType planVFT = GetViewFamilyType(doc, "plan");
					ViewFamilyType rcpVFT = GetViewFamilyType(doc, "rcp");

					foreach (LevelStruct curLevel in levelData)
					{
						Level newLevel = Level.Create(doc, curLevel.LevelElev);
						newLevel.Name = curLevel.LevelName;
						levelCounter++;

						ViewPlan curFloorPlan = ViewPlan.Create(doc, planVFT.Id, newLevel.Id);
						ViewPlan curRCP = ViewPlan.Create(doc, rcpVFT.Id, newLevel.Id);

						curRCP.Name = curRCP.Name + " RCP";
					}

					FilteredElementCollector collector = GetTitleBlock(doc);

					foreach (SheetStruct curSheet in sheetData)
					{
						ViewSheet newSheet = ViewSheet.Create(doc, collector.FirstElementId());
						newSheet.SheetNumber = curSheet.SheetNumber;
						newSheet.Name = curSheet.SheetName;
						SetParameterValue(newSheet, "Drawn By", curSheet.DrawnBy);
						SetParameterValue(newSheet, "Checked By", curSheet.CheckedBy);

						View curView = GetViewByName(doc, curSheet.SheetView);

						if (curView != null)
						{
							Viewport curVP = Viewport.Create(doc, newSheet.Id, curView.Id, new XYZ(0.5, 0.5, 0));
						}

						sheetCounter++;
					}

					t.Commit();
				}

			}

			catch (Exception ex)
			{
				Debug.Print(ex.Message);
			}

			TaskDialog.Show("Complete", "Created " + levelCounter.ToString() + " levels.");
			TaskDialog.Show("Complete", "Created " + sheetCounter.ToString() + " sheets.");

			return Result.Succeeded;
		}

		private static FilteredElementCollector GetTitleBlock(Document doc)
		{
			FilteredElementCollector collector = new FilteredElementCollector(doc);
			collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
			collector.WhereElementIsElementType();
			return collector;
		}

		private View GetViewByName(Document doc, string viewName)
		{
			FilteredElementCollector collector = new FilteredElementCollector(doc);
			collector.OfCategory(BuiltInCategory.OST_Viewers);

			foreach (View curView in collector)
			{
				if (curView.Name == viewName)
					return curView;
			}

			return null;
		}

		private void SetParameterValue(ViewSheet newSheet, string paramName, string paramValue)
		{
			foreach (Parameter curParam in newSheet.Parameters)
			{
				if (curParam.Definition.Name == paramName)
				{
					curParam.Set(paramValue);
				}
			}
		}

		private ViewFamilyType GetViewFamilyType(Document doc, string type)
		{
			FilteredElementCollector collector = new FilteredElementCollector(doc);
			collector.OfClass(typeof(ViewFamilyType));

			foreach (ViewFamilyType vft in collector)
			{
				if (vft.ViewFamily == ViewFamily.FloorPlan && type == "plan")
				{
					return vft;
				}
				else if (vft.ViewFamily == ViewFamily.CeilingPlan && type == "rcp")
				{
					return vft;
				}
			}

			return null;
		}

		private List<SheetStruct> GetSheetDataFromExcel(Excel.Worksheet excelWs)
		{
			List<SheetStruct> returnList = new List<SheetStruct>();
			Excel.Range excelRange1 = excelWs.UsedRange;

			int rowCount1 = excelRange1.Rows.Count;


			for (int i = 2; i <= rowCount1; i++)
			{
				Excel.Range data1 = excelWs.Cells[i, 1];
				Excel.Range data2 = excelWs.Cells[i, 2];
				Excel.Range data3 = excelWs.Cells[i, 3];
				Excel.Range data4 = excelWs.Cells[i, 4];
				Excel.Range data5 = excelWs.Cells[i, 5];

				SheetStruct curSheet = new SheetStruct();
				curSheet.SheetNumber = data1.Value.ToString();
				curSheet.SheetName = data2.Value.ToString();
				curSheet.SheetView = data3.Value.ToString();
				curSheet.DrawnBy = data4.Value;
				curSheet.CheckedBy = data5.Value;

				returnList.Add(curSheet);
			}

			return returnList;
		}

		private List<LevelStruct> GetLevelDataFromExcel(Excel.Worksheet excelWs1)
		{

			// Loop through all the data and put it into a list

			List<LevelStruct> returnList = new List<LevelStruct>();
			Excel.Range excelRange1 = excelWs1.UsedRange;

			int rowCount1 = excelRange1.Rows.Count;


			for (int i = 2; i <= rowCount1; i++)
			{
				Excel.Range levelData1 = excelWs1.Cells[i, 1];
				Excel.Range levelData2 = excelWs1.Cells[i, 2];

				string levelName = levelData1.Value.ToString();
				double levelElev = levelData2.Value;

				LevelStruct curLevel = new LevelStruct(levelName, levelElev);
				returnList.Add(curLevel);
			}

			return returnList;
		}

		private Excel.Worksheet GetExcelWorksheetByName(Excel.Workbook curWb, string wsName)
		{
			foreach (Excel.Worksheet ws in curWb.Worksheets)
			{
				if (ws.Name == wsName)
				{
					return ws;
				}
			}

			return null;
		}

		private struct LevelStruct
		{
			public string LevelName;
			public double LevelElev;

			// create a constructor
			public LevelStruct(string name, double elev)
			{
				LevelName = name;
				LevelElev = elev;
			}
		}

		private struct SheetStruct
		{
			public string SheetNumber;
			public string SheetName;
			public string SheetView;
			public string DrawnBy;
			public string CheckedBy;

			public SheetStruct(string number, string name, string view, string db, string cb)
			{
				SheetNumber = number;
				SheetName = name;
				SheetView = view;
				DrawnBy = db;
				CheckedBy = cb;

			}
		}
	}
}