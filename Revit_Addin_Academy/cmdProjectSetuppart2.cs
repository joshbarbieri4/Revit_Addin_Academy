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
	public class cmdProjectSetuppart2 : IExternalCommand
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
			dialog.Filter = "Excel Files | *.xls; *.xlsx";  // What type of file you want the user to select
						
			string[] filePaths;

			if (dialog.ShowDialog() == Forms.DialogResult.OK)
			{
				filePaths = dialog.FileNames;
			}

			string excelFile = dialog.FileName;

			Excel.Application excelApp = new.Excel.Application(); // Creating an instance of Excel or getting the application
			Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile); // Open Excel to get the Workbook or the particular file above
			Excel.Worksheet excelWs = excelWb.Worksheets.Item[1]; //Now getting the specific Worksheet of the file name

			Excel.Range excelRng = excelWs.UsedRange; // Creating a Range or Selection of cells in use
			int rowCount = excelRng.Rows.Count; //

			return Result.Succeeded;
		}
	}
}
