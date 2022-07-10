#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;

#endregion

namespace Revit_Addin_Academy
{
	[Transaction(TransactionMode.Manual)]
	public class Command01Challenge : IExternalCommand
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

			//string text = "Revit Add-In Academy";

			// Variables for Session 01 Challenge

			string filename = doc.PathName;
			string output1 = "Fizz";
			string output2 = "Buzz";
			int range = 100;
			double offset = 0.05;
			double offsetCalc = offset * doc.ActiveView.Scale;

			// Variables to offset text notes

			XYZ curPoint = new XYZ(0, 0, 0);
			XYZ offsetPoint = new XYZ(0, offsetCalc, 0);

			// Collect the first text note to use below
			FilteredElementCollector collector = new FilteredElementCollector(doc);
			collector.OfClass(typeof(TextNoteType));


			//Begin Transaction on Revit Model
			Transaction t = new Transaction(doc, "Create Text Note");
			t.Start();


			// For Loop to see if number is divisible by 3 &/or 5 or both, if not print number to text
		
			for (int i = 1; i <= range; i++)
			{
				if (i % 3 == 0 && i % 5 == 0)
				{
					TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, output1 + output2, collector.FirstElementId());
					curPoint = curPoint.Subtract(offsetPoint);
				}
				if (i % 3 == 0)
				{
					TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, output1, collector.FirstElementId());
					curPoint = curPoint.Subtract(offsetPoint);
				}
				else if (i % 5 == 0)
				{
					TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, output2, collector.FirstElementId());
					curPoint = curPoint.Subtract(offsetPoint);
				}
				else
				{
					TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, i.ToString(), collector.FirstElementId());
					curPoint = curPoint.Subtract(offsetPoint);
				}
			}

			// End Transaction
			t.Commit();
			t.Dispose();

			return Result.Succeeded;
		}
	}
}
