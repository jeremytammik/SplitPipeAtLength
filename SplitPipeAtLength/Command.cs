#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

#endregion

namespace SplitPipeAtLength
{
  [Transaction(TransactionMode.Manual)]
  public class Command : IExternalCommand
  {
    class PipeSelectionFilter : ISelectionFilter
    {
      /// <summary>
      /// Return true if the element is a pipe 
      /// and the family type name contains "Dyka".
      /// </summary>
      public bool AllowElement(Element e)
      {
        bool rc = false;
        if (e is Pipe)
        {
          string familytypename 
            = e.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
              .AsValueString();

          rc = familytypename.Contains("Dyka");
        }
        return rc;
      }

      public bool AllowReference(Reference r, XYZ p)
      {
        return true;
      }
    }

    //SplitPipe void, need a doc and a pipe as element
    public void SplitPipe(Document doc, Pipe pipe)
    {

      // Get the curve of the pipe
      Line line = (pipe.Location as LocationCurve).Curve as Line;
      //Startpoint
      XYZ startPoint = line.GetEndPoint(0);
      //Endpoint
      XYZ endPoint = line.GetEndPoint(1);

      // Initialize variables for the loop
      XYZ currentPoint = startPoint;
      //Length line
      double remainingLength = startPoint.DistanceTo(endPoint);
      // i for sequence
      int i = 0;

      using (Transaction tx = new Transaction(doc))
      {
        tx.Start("Opdelen leiding");
        //While remainlength > 5000mm
        while (remainingLength > UnitUtils.ConvertToInternalUnits(5000, DisplayUnitType.DUT_MILLIMETERS))
        {
          i++;


          //Check size of the pipe
          string size = pipe.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString();
          //Check name of the pipe, distance can be different (1mm or 1.5mm), PE is different for each size
          string name = pipe.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString();

          //a = distance between connector and splitpoint
          double a;

          //If name and sizes are contains this, than distance = a
          if (name.Contains("NLRS_52_PI_PVC U3")
              && (size.Contains("50")
              || size.Contains("75")
              || size.Contains("110")
              || size.Contains("125")))
          {
            a = 1;
          }
          //If name and sizes are contains this, than distance = a
          else if (name.Contains("NLRS_52_PI_PP")
              && (size.Contains("50")
              || size.Contains("90")
              || size.Contains("110")
              || size.Contains("125")))
          {
            a = 1;
          }
          //If name and sizes are contains this, than distance = a
          else if (name.Contains("NLRS_52_PI_PVC DykaSono")
              && (size.Contains("50")
              || size.Contains("75")
              || size.Contains("125")))
          {
            a = 1;
          }
          else if (name.Contains("NLRS_52_PI_PVC DykaSono")
              && (size.Contains("110")))
          {
            a = 1.5;
          }
          //If name and sizes are contains this, than distance = a
          else if (name.Contains("NLRS_52_PI_UN PVC SN") && size.Contains("125"))
          {
            a = 1;
          }

          //If name and sizes are contains this, than distance = a
          else if (name.Contains("NLRS_52_PI_UN PVC SN") && size.Contains("110") || size.Contains("160"))
          {
            a = 1.5;
          }

          //If name and sizes are contains this, than distance = a
          else if (name.Contains("NLRS_52_PI_PE") && size.Contains("50"))
          {
            a = -1.7;
          }
          //If name and sizes are contains this, than distance = a
          else if (name.Contains("NLRS_52_PI_PE") && size.Contains("75"))
          {
            a = -2.5;
          }
          //If name and sizes are contains this, than distance = a
          else if (name.Contains("NLRS_52_PI_PE") && size.Contains("90")
              || size.Contains("110"))
          {
            a = -3.7;
          }

          //If name and sizes are contains this, than distance = a
          else if (name.Contains("NLRS_52_PI_PE") && size.Contains("125"))
          {
            a = -4.2;
          }

          //else distance = a
          else
          {
            a = 1.5;
          }

          //Use a remainder to check the restlength
          double rest = remainingLength % UnitUtils.ConvertToInternalUnits(5000, DisplayUnitType.DUT_MILLIMETERS);
          double segmentLength;

          //Check of rest length is greater than 100mm. Than use the splitlength (5000) + connector distance
          if (rest > UnitUtils.ConvertToInternalUnits(100, DisplayUnitType.DUT_MILLIMETERS))
          {
            //First item need to be length + distance, after that it need to be length + (2* distance)
            if (i == 1)
            {
              segmentLength = UnitUtils.ConvertToInternalUnits(5000 + a, DisplayUnitType.DUT_MILLIMETERS);
            }
            else
            {
              segmentLength = UnitUtils.ConvertToInternalUnits(5000 + (2 * a), DisplayUnitType.DUT_MILLIMETERS);
            }
          }
          //else split length in 4000mm + connector distance 
          else
          {
            //First item need to be length + distance, after that it need to be length + (2* distance)
            if (i == 1)
            {
              segmentLength = UnitUtils.ConvertToInternalUnits(4000 + a, DisplayUnitType.DUT_MILLIMETERS);
            }
            else
            {
              segmentLength = UnitUtils.ConvertToInternalUnits(4000 + (2 * a), DisplayUnitType.DUT_MILLIMETERS);
            }
          }



          // Calculate the length of the current segment
          double currentSegmentLength = Math.Min(segmentLength, remainingLength);
          // Create the points
          XYZ nextPoint = currentPoint + (line.Direction.Normalize() * currentSegmentLength);

          // Split the pipe at the current segment
          ElementId individualElementId = PlumbingUtils.BreakCurve(doc, pipe.Id, nextPoint);
          //Retrieve the element
          Pipe individualElement = doc.GetElement(individualElementId) as Pipe;

          // Connect the two segments
          foreach (Connector connector in pipe.ConnectorManager.Connectors)
          {
            //Origin connector position
            XYZ connectorOriginPoint = connector.Origin;
            //Create connectorset
            ConnectorSet newPipeConnectorSet = individualElement.ConnectorManager.Connectors;

            foreach (Connector newConnector in newPipeConnectorSet)
            {
              // if connectors distance < 0.01 than create NewUnionFitting
              if (newConnector.Origin.DistanceTo(connectorOriginPoint) < 0.01)
              {
                try
                {
                  doc.Create.NewUnionFitting(connector, newConnector);
                }
                catch { }
                break;
              }
            }
          }

          // Update the variables for the next iteration
          currentPoint = nextPoint;
          remainingLength -= currentSegmentLength;
        }
        tx.Commit();
      }

    }


    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Application app = uiapp.Application;
      Document doc = uidoc.Document;

      // Select a pipe and use the pipeselection filter
      ISelectionFilter pipeFilter = new PipeSelectionFilter();
      try
      {
        Reference[] references = (Reference[])uidoc.Selection.PickObjects(ObjectType.Element, pipeFilter, "Selecteer pipes").ToArray();
        foreach (Reference reference in references)
        {
          Pipe pipes = doc.GetElement(reference.ElementId) as Pipe;

          SplitPipe(doc, pipes);

        }
        return Result.Succeeded;
      }
      catch
      {
        // Do nothing if the selection is null                
        return Result.Cancelled;
      }
    }
  }

  public class Command2 : IExternalCommand
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

      // Access current selection

      Selection sel = uidoc.Selection;

      // Retrieve elements from database

      FilteredElementCollector col
        = new FilteredElementCollector(doc)
          .WhereElementIsNotElementType()
          .OfCategory(BuiltInCategory.INVALID)
          .OfClass(typeof(Wall));

      // Filtered element collector is iterable

      foreach (Element e in col)
      {
        Debug.Print(e.Name);
      }

      // Modify document within a transaction

      using (Transaction tx = new Transaction(doc))
      {
        tx.Start("Transaction Name");
        tx.Commit();
      }

      return Result.Succeeded;
    }
  }
}
