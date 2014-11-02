using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Collections;
using System.Diagnostics;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Structure;
using System.Security.Util;
using System.IO;
using Autodesk.Revit.DB.Events;
using System.Reflection;
using System.Globalization;
using System.Resources;
using Autodesk.Revit.DB.Architecture;

namespace DataToBim
{
  [Transaction( TransactionMode.Manual )]
  [Regeneration( RegenerationOption.Manual )]
  [Journaling( JournalingMode.NoCommandData )]

  public class MainClass : IExternalCommand
  {
    public Autodesk.Revit.UI.Result Execute( ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements )
    {
      const string _data_folder = "C:/a/vs/DataToBim2/DataToBim/TextForPartialMap/";
      try
      {
        Document doc = commandData.Application.ActiveUIDocument.Document;

        #region Save the file


        ////Edited by Mohammad
        //SaveFileDialog saveFileDialog  = new SaveFileDialog();
        //saveFileDialog.Title = "To avoid loosing the data, we recommend saving the file!";
        //saveFileDialog.Filter = "Revit Project Files | *.rvt";
        //saveFileDialog.DefaultExt = "rvt";
        //SaveAsOptions opt = new SaveAsOptions();
        //opt.OverwriteExistingFile = true;
        //bool savefile = true;
        string fileName = "";
        //if (saveFileDialog.ShowDialog() == DialogResult.OK)
        //{
        //    fileName = saveFileDialog.FileName;
        //}
        //else
        //{
        //    savefile = false;
        //}
        //if (savefile)
        //{
        //    doc.SaveAs(fileName, opt);
        //}
        #endregion

        ///****************************************Change the Paths
        List<Building> buildings = EnvironmentalComponents.LoadBuildings( _data_folder + "Buildings.txt" );
        List<Road> roads = EnvironmentalComponents.LoadRoads( _data_folder + "Roads.txt" );
        List<Contour> allContours = EnvironmentalComponents.LoadContours( _data_folder + "Contour.txt" );



        #region timer set up
        Stopwatch timer = Stopwatch.StartNew();
        timer.Reset();
        StringBuilder report = new StringBuilder();
        #endregion

        #region Getting Topography
        /*
                */
        timer.Start();
        List<List<XYZ>> contours = new List<List<XYZ>>();
        foreach( Contour cntr in allContours )
        {
          contours.Add( cntr.vertices );
        }
        DataToTopography getTopo = new DataToTopography( doc, contours, 10, 5 );
        TopographySurface topoSurface = getTopo.Topography;
        timer.Stop();
        report.AppendLine( "Topography Information:" );
        report.AppendLine( timer.Elapsed.TotalSeconds.ToString() + " seconds took to process the contour lines and get the points." );
        report.AppendLine( topoSurface.GetPoints().Count.ToString() + " points exist in the topography." );
        report.AppendLine( getTopo.NumberOfFailedPoints.ToString() + " points were located on the top of each other!" );
        #endregion

        #region Save the file

        //Modified by Mohammad
        //if (savefile)
        //{
        //    doc.SaveAs(fileName, opt);
        //}
        #endregion

        #region Getting Zone-based subregions
        //timer.Reset();
        //timer.Start();
        //List<ProcessPolygon> roadsPoints = new List<ProcessPolygon>();
        //foreach (Road rd in roads)
        //{
        //    roadsPoints.Add(new ProcessPolygon(rd.vertices));
        //}

        //DataToSiteSubRegion getSubregions = new DataToSiteSubRegion(doc, roadsPoints, 300, topoSurface.Id, fileName);
        //getSubregions.AssignColor(new Color(230, 190, 138));
        //timer.Stop();
        //report.AppendLine("");
        //report.AppendLine("Site subregion geration information:");
        //report.AppendLine(timer.Elapsed.TotalSeconds.ToString() + " seconds was needed to create the reagions!");
        //report.AppendLine(getSubregions.FailedAttempts.ToString() + " times failed to create subregions in zones");

        //#endregion


        //#region Save the file
        //if (savefile)
        //{
        //    doc.SaveAs(fileName, opt);
        //}
        #endregion

        #region getting the buildings


        timer.Reset();
        timer.Start();
        DataToBuilding getBldgs = new DataToBuilding( doc, buildings, getTopo.Topography.Id, fileName );
        timer.Stop();
        report.AppendLine( "" );
        report.AppendLine( "Building & building pad geration information:" );
        report.AppendLine( timer.Elapsed.TotalSeconds.ToString() + " seconds was needed to create the building!" );
        report.AppendLine( getBldgs.FailedAttemptsToCreateBuildings.ToString() + " times failed to generate buildings!" );
        report.AppendLine( getBldgs.FailedAttemptsToCreateBuildingPads.ToString() + " times failed to generate building pads!" );

        #endregion

        #region Save the file
        //Modified by Mohammad
        //if (savefile)
        //{
        //    doc.SaveAs(fileName, opt);
        //}

        #endregion

        #region Creating a report
        TaskDialog.Show( "Report of process", report.ToString() );
        #endregion



      }
      catch( Exception exception )
      {
        message = exception.Message;
        return Autodesk.Revit.UI.Result.Failed;
      }
      return Autodesk.Revit.UI.Result.Succeeded;
    }
  }

  //this is an implementation for IFailuresPreprocessor interface to swallow warnings
  public class WarningSwallower : IFailuresPreprocessor
  {
    public FailureProcessingResult PreprocessFailures( FailuresAccessor a )
    {
      // inside event handler, get all warnings
      IList<FailureMessageAccessor> failures = a.GetFailureMessages();
      foreach( FailureMessageAccessor f in failures )
      {
        a.DeleteAllWarnings();
      }
      return FailureProcessingResult.Continue;
    }
  }
}
