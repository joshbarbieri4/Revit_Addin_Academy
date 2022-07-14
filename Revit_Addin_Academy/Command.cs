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

#endregion

namespace Revit_Addin_Academy
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

			string excelFile = @"L:\z_BIM-CAD-Studies\6_REVIT ADDIN ACADEMY\02_Session 02 - Challenge_Download\Session02_Challenge-JB.xlsx";

			Excel.Application excelApp = new Excel.Application(); // Creating an instance of Excel or getting the application
			Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile); // Open Excel to get the workbook or the particular file above
			Excel.Worksheet excelWs = excelWb.Worksheets.Item[1]; // Now getting the specific Worksheet of the file above

			Excel.Range excelRng = excelWs.UsedRange; // Creating a Range or selection of cells in use
			int rowCount = excelRng.Rows.Count; // Reports us how many rows there are in the excel file

			// do some stuff in Excel

			List<string[]> dataList = new List<string[]>(); // Group all arrays together

			for (int i = 1; i < rowCount; i++) // Loop through each row, one at time, and get specific cells
			{
				Excel.Range cell1 = excelWs.Cells[i, 1]; // X-Y of a specific cell
				Excel.Range cell2 = excelWs.Cells[i, 2];

				string data1 = cell1.Value.ToString(); // get the values of the specific cells above and putting into a variable
				string data2 = cell2.Value.ToString();

				string[] dataArray = new string[2]; // creating an array [like a box with two spots that are empty]
				dataArray[0] = data1; // store data in array
				dataArray[1] = data2;

				dataList.Add(dataArray);

			}

			using(Transaction t = new Transaction(doc))
			{
				t.Start("Create some Revit Stuff");

				Level curLevel = Level.Create(doc, 100);

				FilteredElementCollector collector = new FilteredElementCollector(doc);
				collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
				collector.WhereElementIsElementType();

				ViewSheet curSheet = ViewSheet.Create(doc, collector.FirstElementId());
				curSheet.SheetNumber = "aaaa";
				curSheet.Name = "New Sheet";
			

				t.Commit();
			}



			// Opens Excel and Close Excel and avoids too many instances of Excel Running in the background.
			excelWb.Close();
			excelApp.Quit();

			return Result.Succeeded;
		}
	}
}
