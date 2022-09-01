#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Architecture;

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
			// General Notes:
			// In C#, there might be two or more methods in a class with the same name but different numbers, types, and order of parameters, it is called method overloading.
			// CurveElement is generic catch all of all types of line based elements for example model line/detail line/area line
			// casting is a way of explicity informing the complier that you intend to make the conversion of the selection.
			// Rooms is part of the Architecture library so make sure and include using Autodesk.Revit.DB.Architecture

			UIApplication uiapp = commandData.Application;  // Is our session of Revit runnig
			UIDocument uidoc = uiapp.ActiveUIDocument;      // Is the document that is running in instances of Revit above			

			Application app = uiapp.Application;            // behind the scenes of the application
			Document doc = uidoc.Document;                  // getting the Revit database

			// I - indicates interface list
			IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("Select some elements:");
			List<CurveElement> curveList = new List<CurveElement>();


			WallType curWallType = GetWallTypeByName(doc, @"Generic - 8""");
			Level curLevel = GetLevelByName(doc, "Level 1");

			MEPSystemType curSystemType = GetSystemTypeByName(doc, "Domestic Hot Water");
			PipeType curPipeType = GetPipeTypeName(doc, "Default");


			using(Transaction t = new Transaction(doc))
			{
				t.Start("Create Revit Stuff");

				// filter the selction of elements that are picked
				// Element is the base class of Revit
				foreach (Element element in pickList)
				{
					// The word is asking if this particular element is part of this class
					if (element is CurveElement)
					{
						// Translating the element over to CurveElement
						CurveElement curve = (CurveElement)element; // casting a curve Element
						CurveElement curve2 = element as CurveElement;

						curveList.Add(curve);

						// look at the properties of the linestyles
						GraphicsStyle curGS = curve.LineStyle as GraphicsStyle; // get name of linestyle
						Curve curCurve = curve.GeometryCurve; // get geometryo of the line
						XYZ startPoint = curCurve.GetEndPoint(0); // get start and end point of the line
						XYZ endPoint = curCurve.GetEndPoint(1);

						// create a wall
						// Wall newWall = Wall.Create(doc, curCurve, curWallType.Id, curLevel.Id, 15, 0, false, false);

						Pipe newPip = Pipe.Create(
							doc,
							curSystemType.Id,
							curPipeType.Id,
							curLevel.Id,
							startPoint,
							endPoint
							);


						Debug.Print(curGS.Name);
					}
				}

				t.Commit();
			}

			
			TaskDialog.Show("complete", curveList.Count.ToString());
			return Result.Succeeded;
		}

		private WallType GetWallTypeByName(Document doc, string wallTypeName)
		{
			FilteredElementCollector collector = new FilteredElementCollector(doc);
			collector.OfClass(typeof(WallType));
			{
				foreach(Element curElem in collector)
				{
					WallType wallType = curElem as WallType; // cast current Element into WallType

					if(wallType.Name == wallTypeName)
						return wallType;
				}

				return null;
			}
		}

		private Level GetLevelByName(Document doc, string levelName)
		{
			FilteredElementCollector collector = new FilteredElementCollector(doc);
			collector.OfClass(typeof(Level));
			{
				foreach (Element curElem in collector)
				{
					Level level = curElem as Level; // cast current Element into Level

					if (level.Name == levelName)
						return level;
				}

				return null;
			}
		}

		private MEPSystemType GetSystemTypeByName (Document doc, string typeName)
		{
			FilteredElementCollector collector = new FilteredElementCollector(doc);
			collector.OfClass(typeof(MEPSystemType));
			{
				foreach (Element curElem in collector)
				{
					MEPSystemType curType = curElem as MEPSystemType; // cast current Element into MEPSystemType

					if (curType.Name == typeName)
						return curType;
				}

				return null;
			}
		}

		private PipeType GetPipeTypeName(Document doc, string typeName)
		{
			FilteredElementCollector collector = new FilteredElementCollector(doc);
			collector.OfClass(typeof(PipeType));
			{
				foreach (Element curElem in collector)
				{
					PipeType curType = curElem as PipeType; // cast current Element into PipeType

					if (curType.Name == typeName)
						return curType;
				}

				return null;
			}
		}

	}
}
