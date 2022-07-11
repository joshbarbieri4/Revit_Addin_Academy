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
			
			// Variables for Session 01 Challenge
			string filename = doc.PathName;			
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
			using(Transaction t = new Transaction(doc))
			{
				t.Start("FizzBuzz");

				// For Loop to see if number is divisible by 3 &/or 5 or both, if not print number to text
				for (int i = 1; i <= range; i++)
				{

					string result = CheckFizzBuzz(i);

					CreateTextNote(doc, result, curPoint, collector.FirstElementId());
					curPoint = curPoint.Subtract(offsetPoint);					
				}

				// End Transaction
				t.Commit();
			}
			
			return Result.Succeeded;
		}				
		internal void CreateTextNote(Document doc, string text, XYZ curPoint, ElementId id)
		{
			TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, text, id);
		}

		internal string CheckFizzBuzz(int number)
		{
			string result = "";

			if (number % 3 == 0)
			{
				result = "Fizz";
			}

			if (number % 5 == 0)
			{
				result = result + "Buzz";
			}

			if (number % 3 != 0 && number % 5 != 0)
			{
				result = number.ToString();
			}

			Debug.Print(result);

			return result;
		}
	}
}