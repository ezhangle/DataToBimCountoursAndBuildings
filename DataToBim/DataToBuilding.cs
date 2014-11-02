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
using View = Autodesk.Revit.DB.View;

namespace DataToBim
{
  class DataToBuilding
  {
    public List<ProcessPolygon> BuildingFootPrints;
    public Document _doc;
    public readonly List<FamilyInstance> BuildingMasses = new List<FamilyInstance>();
    private View3D RayTracerView;
    public ElementId TopographyID;
    private readonly string FamilyTemplateAddress;
    public readonly BuildingPadType buildingPadType;
    public readonly List<Autodesk.Revit.DB.Architecture.BuildingPad> BuildingPads = new List<Autodesk.Revit.DB.Architecture.BuildingPad>();
    public int FailedAttemptsToCreateBuildings;
    public int FailedAttemptsToCreateBuildingPads;
    private SaveAsOptions opt = new SaveAsOptions();
    private string DocumentFileAddress;
    public DataToBuilding( Document doc, List<Building> buildings, ElementId topographyID, string documentFileAddress = "" )
    {
      this.DocumentFileAddress = documentFileAddress;
      this.opt.OverwriteExistingFile = true;
      List<double> bldgHeight = new List<double>();
      _doc = doc; this.TopographyID = topographyID;
      using( Transaction createBuildingPadType = new Transaction( _doc ) )
      {
        createBuildingPadType.Start( "Create Building Pad Type" );
        this.buildingPadType = BuildingPadType.CreateDefault( _doc );
        createBuildingPadType.Commit();
      }
      this.FailedAttemptsToCreateBuildingPads = 0;
      this.FailedAttemptsToCreateBuildings = 0;

      #region pre-process the contour lines
      List<ProcessPolygon> ProcessedFootPrints = new List<ProcessPolygon>();
      this.RayTracerView = this.createView3d();
      foreach( Building building in buildings )
      {
        ProcessPolygon processedContour = new ProcessPolygon( building.vertices );
        processedContour.Flatten();
        processedContour.RemoveIdenticalPoints();
        processedContour.RemoveClosePoints( .5 );
        processedContour.RemoveCollinearity();
        processedContour.ForceToFixList();

        bldgHeight.Add( building.height );
        ProcessedFootPrints.Add( processedContour );
      }
      this.BuildingFootPrints = ProcessedFootPrints;
      #endregion

      #region Create building Masses ad independent families

      this.FamilyTemplateAddress = this.familyTemplateAddress();
      List<double> elevations = new List<double>();
      for( int i = 0; i < this.BuildingFootPrints.Count; i++ )
      {
        elevations.Add( this.FindBuildingElevation( this.BuildingFootPrints[i] ) );
      }
      //Finding a folder to save the buildings
      FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
      DialogResult result = folderBrowser.ShowDialog();
      string folder = "";
      if( result == DialogResult.OK )
      {
        folder = folderBrowser.SelectedPath;
      }
      //creating extrusions and saving them
      List<XYZ> translations = new List<XYZ>();
      List<FamilySymbol> symbols = new List<FamilySymbol>();
      List<int> footPrintIndices = new List<int>();

      for( int i = 0; i < 5; i++ )
      {
        if( elevations[i] != double.NaN && elevations[i] != -1 )
        {
          string path = folder + @"\" + i.ToString() + ".rfa";
          XYZ translation = XYZ.Zero;
          foreach( XYZ item in this.BuildingFootPrints[i].ProcessedPolygon )
          {
            translation += item;
          }
          translation /= this.BuildingFootPrints[i].ProcessedPolygon.Count;

          bool formCreated = CreateFamilyFile( this.BuildingFootPrints[i], bldgHeight[i], path, translation );
          if( formCreated )
          {
            using( Transaction placeFamily = new Transaction( _doc ) )
            {
              FailureHandlingOptions failOpt = placeFamily.GetFailureHandlingOptions();
              failOpt.SetFailuresPreprocessor( new WarningSwallower() );
              placeFamily.SetFailureHandlingOptions( failOpt );
              placeFamily.Start( "Place a Mass" );
              Family family = null;
              _doc.LoadFamily( path, out family );
              FamilySymbol symbol = null;
              //foreach( FamilySymbol s in family.Symbols ) // 2014
              foreach( ElementId id in family.GetFamilySymbolIds() ) // 2015
              {
                symbol = _doc.GetElement( id ) as FamilySymbol;
                break;
              }
              symbols.Add( symbol );
              XYZ displacement = new XYZ( translation.X, translation.Y, elevations[i] );
              translations.Add( displacement );
              footPrintIndices.Add( i );
              placeFamily.Commit();
            }
          }
          else
          {
            this.BuildingFootPrints[i].Visualize( _doc );
          }
        }
      }
      #endregion

      #region inserting a mass file
      for( int i = 0; i < translations.Count; i++ )
      {
        using( Transaction placeMass = new Transaction( _doc ) )
        {
          FailureHandlingOptions failOpt = placeMass.GetFailureHandlingOptions();
          failOpt.SetFailuresPreprocessor( new WarningSwallower() );
          placeMass.SetFailureHandlingOptions( failOpt );
          placeMass.Start( "Insert Mass" );
          try
          {
            FamilyInstance building = _doc.Create.NewFamilyInstance( XYZ.Zero, symbols[i], StructuralType.NonStructural );
            ElementTransformUtils.MoveElement( _doc, building.Id, translations[i] );
            this.BuildingMasses.Add( building );
          }
          catch( Exception )
          {
            this.FailedAttemptsToCreateBuildings++;
          }
          placeMass.Commit();
        }
      }
      #endregion

      if( this.DocumentFileAddress != "" )
      {
        _doc.SaveAs( this.DocumentFileAddress, this.opt );
      }

      #region creating building pads
      //for (int i = 0; i < footPrintIndices.Count; i++)
      //{
      //    Transaction CreatePads = new Transaction(_doc, "Create Building Pad");
      //    FailureHandlingOptions failOpt = CreatePads.GetFailureHandlingOptions();
      //    failOpt.SetFailuresPreprocessor(new WarningSwallower());
      //    CreatePads.SetFailureHandlingOptions(failOpt);
      //    CreatePads.Start();
      //    try
      //    {
      //        Autodesk.Revit.DB.Architecture.BuildingPad pad = this.CreateBuildingPad(this.BuildingFootPrints[footPrintIndices[i]], elevations[footPrintIndices[i]]);
      //        this.BuildingPads.Add(pad);
      //        if (this.DocumentFileAddress != "")
      //        {
      //            _doc.SaveAs(this.DocumentFileAddress, this.opt);
      //        }
      //    }
      //    catch (Exception)
      //    {
      //        this.FailedAttemptsToCreateBuildingPads++;
      //    }

      //    CreatePads.Commit();
      //}
      #endregion


    }
    // We need to create a 3dview because the when we open a new project file a 3d view does not exist and raytracers need 3dviews
    private View3D createView3d()
    {
      FilteredElementCollector collector0 = new FilteredElementCollector( _doc ).OfClass( typeof( View3D ) );
      foreach( View3D item in collector0 )
      {
        if( !item.IsTemplate )
        {
          return item;
        }
      }
      FilteredElementCollector collector1 = new FilteredElementCollector( _doc );
      collector1 = collector1.OfClass( typeof( ViewFamilyType ) );
      IEnumerable<ViewFamilyType> viewFamilyTypes = from elem in collector1 let vftype = elem as ViewFamilyType where ( vftype.ViewFamily == ViewFamily.ThreeDimensional ) select vftype;
      using( Transaction createView3D = new Transaction( _doc ) )
      {
        createView3D.Start( "Create 3D View" );
        View3D view3D = View3D.CreateIsometric( _doc, viewFamilyTypes.First<ViewFamilyType>().Id );
        if( null != view3D )
        {
          XYZ eye = new XYZ( 10, 10, 10 );
          XYZ up = new XYZ( 0, 1, 1 );
          XYZ forward = new XYZ( 0, 1, -1 );
          ViewOrientation3D viewOrientation3D = new ViewOrientation3D( eye, up, forward );
          view3D.SetOrientation( viewOrientation3D );
          view3D.Name = "RayTracer View";
        }
        createView3D.Commit();
        return view3D;
      }
    }
    // finding building elevation
    public double FindBuildingElevation( ProcessPolygon polygon )
    {
      List<XYZ> heights = new List<XYZ>();
      ReferenceIntersector refIntersector = new ReferenceIntersector( this.TopographyID, FindReferenceTarget.Mesh, this.RayTracerView );
      int counter = 0;
      double h = 0;
      foreach( XYZ item in polygon.ProcessedPolygon )
      {
        ReferenceWithContext referenceWithContext = refIntersector.FindNearest( item, XYZ.BasisZ );
        if( referenceWithContext != null )
        {
          counter++;
          h += referenceWithContext.Proximity;
        }
        else
        {
          return -1;
        }
      }
      return h / counter;
    }
    // find the address of Conceptual Mass Family Templage
    public string familyTemplateAddress()
    {
      IDictionary<string, string> revitPaths = _doc.Application.GetLibraryPaths();
      string path = "";
      if( !revitPaths.TryGetValue( "Imperial Library", out path )
        && !revitPaths.TryGetValue( "Metric Library", out path ) )
      {
        throw new ArgumentException( "Could not find the path! \n" );
      }
      DirectoryInfo lib = new DirectoryInfo( path );
      string s = "";
      try
      {
        DirectoryInfo RVT = lib.Parent.Parent;
        s = Path.Combine( RVT.FullName + @"\Family Templates\English_I\Conceptual Mass\Mass.rft" );
        if( !File.Exists( s ) )
        {
          s = Path.Combine( RVT.FullName + @"\Family Templates\English\Conceptual Mass\Metric Mass.rft" );
        }
        if( !File.Exists( s ) )
        {
          throw new ArgumentException( "Could not find the conceptual mass template! \n" );
        }
      }
      catch( Exception er )
      {
        throw new ArgumentException( er.Message );
      }
      return s;
    }
    //Create family for each building 
    //http://thebuildingcoder.typepad.com/blog/2011/06/creating-and-inserting-an-extrusion-family.html
    private bool CreateFamilyFile( ProcessPolygon polygon, double height, string familyFileName, XYZ translation )
    {
      bool success = true;
      Document FamDoc = null;
      Autodesk.Revit.DB.Form form = null;
      using( Transaction CreateFamily = new Transaction( _doc ) )
      {
        CreateFamily.Start( "Create a new Family" );
        FamDoc = _doc.Application.NewFamilyDocument( this.FamilyTemplateAddress );
        CreateFamily.Commit();
      }
      ReferenceArray refAr = polygon.Get_ReferenceArray( FamDoc, translation );
      using( Transaction CreateExtrusion = new Transaction( FamDoc ) )
      {
        FailureHandlingOptions failOpt = CreateExtrusion.GetFailureHandlingOptions();
        failOpt.SetFailuresPreprocessor( new WarningSwallower() );
        CreateExtrusion.SetFailureHandlingOptions( failOpt );
        CreateExtrusion.Start( "Create Extrusion" );
        /*Mohammad took this out of try */
        try
        {
          form = FamDoc.FamilyCreate.NewExtrusionForm( true, refAr, height * XYZ.BasisZ );
        }
        catch( Exception )
        {
          this.FailedAttemptsToCreateBuildings++;
          success = false;
        }
        CreateExtrusion.Commit();
      }

      //Added by Mohammad
      using( Transaction AddParamTrans = new Transaction( FamDoc ) )
      {
        AddParamTrans.Start( "Add Parameter" );
        Autodesk.Revit.ApplicationServices.Application app = FamDoc.Application;
        View3D view3d = createView3d();

        Dimension windowInsetDimension = null;
        FaceArray faces = new FaceArray();
        if( form.IsSolid )
        {
          Options options = new Options();
          options.ComputeReferences = true;
          //options.View = new Autodesk.Revit.DB.View();
          //GeometryObjectArray geoArr = extrusion.get_Geometry(options).Objects;
          IEnumerator<GeometryObject> Objects = form.get_Geometry( options ).GetEnumerator();
          //foreach (GeometryObject geoObj in geoArr)
          while( Objects.MoveNext() )
          {
            GeometryObject geoObj = Objects.Current;

            if( geoObj is Solid )
            {
              Solid s = geoObj as Solid;
              foreach( Face fc in s.Faces )
              {
                //MessageBox.Show(fc.ComputeNormal(new UV(0, 0)).X.ToString() + "/n" + fc.ComputeNormal(new UV(0, 0)).Y.ToString() + "/n" + fc.ComputeNormal(new UV(0, 0)).Z.ToString());
                if( Math.Round( ( fc.ComputeNormal( new UV( 0, 0 ) ) ).Z ) == 1 || Math.Round( ( fc.ComputeNormal( new UV( 0, 0 ) ) ).Z ) == -1 )
                {
                  faces.Append( fc );
                }
              }
              //**************************************************************************************************************
              //****************************Here is the Error **********************************************************************
              //************************************************************************************************************************
              //windowInsetDimension = AddDimension( FamDoc, view3d, faces.get_Item( 0 ), faces.get_Item( 1 ) );
              View viewElevation = new FilteredElementCollector( FamDoc ).OfClass( typeof( View ) ).Cast<View>().Where<View>( v => ViewType.Elevation == v.ViewType ).FirstOrDefault<View>();
              windowInsetDimension = AddDimension( FamDoc, viewElevation, faces.get_Item( 0 ), faces.get_Item( 1 ) );
            }
          }
        }

        //Test for creating dimension
        #region two lines creating dimenstion
        //// first create two lines
        //XYZ pt1 = new XYZ(5, 0, 5);
        //XYZ pt2 = new XYZ(5, 0, 10);
        //Line line = Line.CreateBound(pt1, pt2);
        //Plane plane = app.Create.NewPlane(pt1.CrossProduct(pt2), pt2);

        //SketchPlane skplane = SketchPlane.Create(FamDoc, plane);

        //ModelCurve modelcurve1 = FamDoc.FamilyCreate.NewModelCurve(line, skplane);

        //pt1 = new XYZ(10, 0, 5);
        //pt2 = new XYZ(10, 0, 10);
        //line = Line.CreateBound(pt1, pt2);
        //plane = app.Create.NewPlane(pt1.CrossProduct(pt2), pt2);

        //skplane = SketchPlane.Create(FamDoc, plane);

        //ModelCurve modelcurve2 = FamDoc.FamilyCreate.NewModelCurve(line, skplane);




        //// now create a linear dimension between them
        //ReferenceArray ra = new ReferenceArray();
        //ra.Append(modelcurve1.GeometryCurve.Reference);
        //ra.Append(modelcurve2.GeometryCurve.Reference);

        //pt1 = new XYZ(5, 0, 10);
        //pt2 = new XYZ(10, 0, 10);
        //line = Line.CreateBound(pt1, pt2);


        //Dimension dim = FamDoc.FamilyCreate.NewLinearDimension(view3d, line, ra);
        #endregion

        //creates a prameter named index for each family.
        BuiltInParameterGroup paramGroup = (BuiltInParameterGroup) Enum.Parse( typeof( BuiltInParameterGroup ), "PG_GENERAL" );
        ParameterType paramType = (ParameterType) Enum.Parse( typeof( ParameterType ), "Length" );
        FamilyManager m_manager = FamDoc.FamilyManager;

        FamilyParameter famParam = m_manager.AddParameter( "Height", paramGroup, paramType, true );

        //Set the value for the parameter
        if( m_manager.Types.Size == 0 )
          m_manager.NewType( "Type 1" );

        m_manager.Set( famParam, height );

        //connects dimension to lable called with
        //windowInsetDimension.FamilyLabel = famParam;

        AddParamTrans.Commit();
      }


      SaveAsOptions opt = new SaveAsOptions();
      opt.OverwriteExistingFile = true;
      FamDoc.SaveAs( familyFileName, opt );
      FamDoc.Close( false );
      return success;
    }
    //create building pads
    public Autodesk.Revit.DB.Architecture.BuildingPad CreateBuildingPad( ProcessPolygon polygon, double elevation )
    {
      Level level = _doc.Create.NewLevel( elevation );
      List<CurveLoop> listOfCurveLoops = new List<CurveLoop>();
      listOfCurveLoops.Add( polygon.Get_CurveLoop() );
      return Autodesk.Revit.DB.Architecture.BuildingPad.Create( _doc, this.buildingPadType.Id, level.Id, listOfCurveLoops );
    }

    public Dimension AddDimension( Document doc, Autodesk.Revit.DB.View view, Face face1, Face face2 )
    {
      Dimension dim;
      Autodesk.Revit.DB.XYZ startPoint = new Autodesk.Revit.DB.XYZ();
      Autodesk.Revit.DB.XYZ endPoint = new Autodesk.Revit.DB.XYZ();
      Line line;
      Reference ref1;
      Reference ref2;
      ReferenceArray refArray = new ReferenceArray();
      PlanarFace pFace1 = face1 as PlanarFace;
      ref1 = pFace1.Reference;
      PlanarFace pFace2 = face2 as PlanarFace;
      ref2 = pFace2.Reference;
      if( null != ref1 && null != ref2 )
      {
        refArray.Append( ref1 );
        refArray.Append( ref2 );
      }
      startPoint = pFace1.Origin;
      endPoint = new Autodesk.Revit.DB.XYZ( startPoint.X, startPoint.Y, pFace2.Origin.Z );
      SubTransaction subTransaction = new SubTransaction( doc );
      subTransaction.Start();
      line = Line.CreateBound( startPoint, endPoint );
      dim = doc.FamilyCreate.NewDimension( view, line, refArray );
      subTransaction.Commit();
      return dim;
    }
  }
}
