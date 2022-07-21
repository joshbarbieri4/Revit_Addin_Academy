#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Forms = System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

#endregion

namespace Revit_Addin_Academy
{
	[Transaction(TransactionMode.Manual)]
	public class Project_Setup_Part2 : IExternalCommand
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
			dialog.Multiselect = true;
			dialog.Filter = "Excel Files | *.xls; *.xlsx"; // What the user sees & specifiy what type of file the user can select

			string excelFile = "";
			
			if(dialog.ShowDialog() == Forms.DialogResult.OK) // opens dialog box, and user clicks ok do something - captures what the user does
			{
				excelFile = dialog.FileName;				
			}

			// open Excel
			Excel.Application excelApp = new Excel.Application(); // created a variable that holds the application and opens it
			Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);

			Excel.Worksheet excelWs1 = excelWb.Worksheets.Item[1];
			Excel.Worksheet excelWs2 = excelWb.Worksheets.Item[2];

			Excel.Range excelRange1 = excelWs1.UsedRange;
			Excel.Range excelRange2 = excelWs2.UsedRange;

			int rowCount1 = excelRange1.Rows.Count;
			int rowCount2 = excelRange2.Rows.Count;

			int levelCounter = 0;

			using (Transaction t = new Transaction(doc))
			{
				t.Start("Setup project");

				for (int i = 2; i <= rowCount1; i++)
				{
					Excel.Range levelData1 = excelWs1.Cells[i, 1]; // get level name
					Excel.Range levelData2 = excelWs1.Cells[i, 2]; // get level elevation

					string levelName = levelData1.Value.ToString();
					double levelElev = levelData2.Value;

					try
					{
						Level newLevel = Level.Create(doc, levelElev);
						newLevel.Name = levelName;
						levelCounter++;
					}
					catch (Exception ex)
					{
						Debug.Print(ex.Message);
						throw;
					}										
				}

				t.Commit();

				excelWb.Close();
				excelApp.Quit();

				return Result.Succeeded;
			}
		}
	
	}

}

