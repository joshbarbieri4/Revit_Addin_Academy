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

			string[] filePaths;
			
			if(dialog.ShowDialog() == Forms.DialogResult.OK) // opens dialog box, and user clicks ok do something - captures what the user does
			{
				// filePath = dialog.FileName;
				filePaths = dialog.FileNames;
			}
													
			return Result.Succeeded;
		}

		internal View GetViewByName(Document doc, string ViewName)
		{
			FilteredElementCollector collector = new FilteredElementCollector(doc);
			collector.OfClass(typeof(View));

			foreach(View curView in collector)
			{
				if(curView.Name == ViewName)
				{
					return curView;
				}
			}

			return null;
		}

		internal struct TestStruct
		{ 
			public string Name;
			public int Value;
			public double Value2;

			// creating a constructor
			public TestStruct(string name, int value, double value2)
			{
				Name = name;
				Value = value;
				Value2 = value2;
			}			
		}
	}

}

