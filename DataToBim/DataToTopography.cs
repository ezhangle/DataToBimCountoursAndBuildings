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
//Clipper library for boolean operations of the polygons: http://sourceforge.net/projects/polyclipping/?source=navbar
using ClipperLib;
using Polygon = System.Collections.Generic.List<ClipperLib.IntPoint>;
using Polygons = System.Collections.Generic.List<System.Collections.Generic.List<ClipperLib.IntPoint>>;
using CurveLoops = System.Collections.Generic.List<Autodesk.Revit.DB.CurveLoop>;

namespace DataToBim
{
    class DataToTopography
    {
        public readonly int NumberOfFailedPoints;
        public readonly TopographySurface Topography;
        public DataToTopography(Document document, List<List<XYZ>> contours, double interpolationLength = 10, double proximityTolerance = 5)
        {
            Dictionary<string, XYZ> dict = new Dictionary<string, XYZ>();
            int counter = 0;
            foreach (List<XYZ> cntr in contours)
            {
                ProcessPolygon processedContour = new ProcessPolygon(cntr);
                processedContour.LoadEdgeLengths();
                processedContour.RemoveClosePoints(proximityTolerance);
                processedContour.Interpolate(interpolationLength);
                processedContour.LoadEdgeLengths();
                foreach (XYZ item in processedContour.ProcessedPolygon)
                {
                    string key = item.X.ToString() + item.Y.ToString();
                    try
                    {
                        dict.Add(key, item);
                    }
                    catch (Exception)
                    {
                        //MessageBox.Show(er.Message);
                        counter++;
                    }
                }
            }
            this.NumberOfFailedPoints = counter;
            List<XYZ> pnts = new List<XYZ>();
            foreach (var item in dict.Keys)
            {
                XYZ pn;
                if (dict.TryGetValue(item, out pn))
                {
                    pnts.Add(pn);
                }
            }
            Transaction createTopography = new Transaction(document, "Create Topography");
            createTopography.Start();
            this.Topography = TopographySurface.Create(document, pnts);
            createTopography.Commit();
        }
    }

    class DataToSiteSubRegion
    {
        public double RegionSize;
        public int XNumber;
        public int YNumber;
        public double RegionWidth;
        public double RegionHeight;
        public int FailedAttempts;
        public ElementId TopographySurfaceId;
        public BoundingBoxUV[,]  ReagionsBoundingBoxes;
        public bool[,] CanSiteSubregionBeValid;
        private View3D RayTracerView;
        private Document document;
        private readonly int Exponent;
        private Polygons IntRoads = new Polygons();
        public List<SiteSubRegion> SiteSubRegions = new List<SiteSubRegion>();
        private string DocumentFileAddress;
        private SaveAsOptions opt = new SaveAsOptions();
        //converting all of the lists into polygons
        public DataToSiteSubRegion(Document doc, List<ProcessPolygon> Roads, double regionSize, ElementId topographySurfaceId, string fileAddress = "")
        {
            this.opt.OverwriteExistingFile = true;
            this.DocumentFileAddress = fileAddress;
            this.document = doc;
            this.TopographySurfaceId = topographySurfaceId;
            this.Exponent = 5;
            this.FailedAttempts = 0;

            #region Contour treatments 
            foreach (ProcessPolygon processedContour in Roads)
            {
                processedContour.Flatten();
                processedContour.RemoveIdenticalPoints();
                processedContour.RemoveClosePoints(.5);
                processedContour.RemoveCollinearity(4 * Math.PI / 180);
                //processedContour.ForceToFixList();
            }
            #endregion

            #region Setting the sizes of regions
            this.RegionSize = regionSize;
            double minX = Roads[0].ProcessedPolygon[0].X;
            double maxX = Roads[0].ProcessedPolygon[0].X;
            double minY = Roads[0].ProcessedPolygon[0].Y;
            double maxY = Roads[0].ProcessedPolygon[0].Y;
            foreach (ProcessPolygon ProcessedPolygon in Roads)
            {
                foreach (XYZ item in ProcessedPolygon.ProcessedPolygon)
                {
                    minX = (minX > item.X) ? item.X : minX;
                    maxX = (maxX < item.X) ? item.X : maxX;
                    minY = (minY > item.Y) ? item.Y : minY;
                    maxY = (maxY < item.Y) ? item.Y : maxY;
                }
            }
            this.XNumber = (int)Math.Ceiling((maxX - minX) / this.RegionSize);
            this.YNumber = (int)Math.Ceiling((maxY - minY) / this.RegionSize);
            this.RegionWidth = (maxX - minX) / this.XNumber;
            this.RegionHeight = (maxY - minY) / this.YNumber;
            this.ReagionsBoundingBoxes = new BoundingBoxUV[this.XNumber, this.YNumber];
            this.CanSiteSubregionBeValid = new bool[this.XNumber, this.YNumber];
            for (int i = 0; i < this.XNumber; i++)
            {
                for (int j = 0; j < this.YNumber; j++)
                {
                    double min_u = minX + i*this.RegionWidth;
                    double max_u = min_u + this.RegionWidth;
                    double min_v = minY + j*this.RegionHeight;
                    double max_v = min_v + this.RegionHeight;
                    this.ReagionsBoundingBoxes[i, j] = new BoundingBoxUV(min_u, min_v, max_u, max_v);
                }
            }
            #endregion

            #region Checking to see if subregions are projectable on the topography
            this.RayTracerView = this.createView3d();
            for (int i = 0; i < this.XNumber; i++)
            {
                for (int j = 0; j < this.YNumber; j++)
                {
                    this.CanSiteSubregionBeValid[i, j] = this.CanSubregionBeValid(this.ReagionsBoundingBoxes[i, j]);
                }
            }
            #endregion

            #region creating intPoint polygons of roads
            Polygons initialRoads = new Polygons();
            foreach (ProcessPolygon XYZPolygon in Roads)
            {
                initialRoads.Add(this.XYZList2Polygon(XYZPolygon.ProcessedPolygon, this.Exponent));
            }
            this.IntRoads = this.XOR(initialRoads);
            #endregion

            #region Subtracting buildings' footprints 
            //placeholder: to be completed
            #endregion

            #region Creating the site sub-regions

            for (int i = 0; i < this.XNumber; i++)
            {
                for (int j = 0; j < this.YNumber; j++)
                {
                    if (this.CanSiteSubregionBeValid[i,j])
                    {
                        bool visualizePolygons = false;
                        Polygons initialPolygons = getPolygonsInRegion(i, j, 10);
                        if (initialPolygons.Count > 0)
                        {
                            CurveLoops curveLoops = this.PolygonsToCurveLoops(initialPolygons);
                            if (curveLoops.Count>0)
                            {
                                Transaction createSubRegions = new Transaction(document, "Create subregions in site");
                                FailureHandlingOptions failOpt = createSubRegions.GetFailureHandlingOptions();
                                failOpt.SetFailuresPreprocessor(new WarningSwallower());
                                createSubRegions.SetFailureHandlingOptions(failOpt);
                                createSubRegions.Start();
                                try
                                {
                                    SiteSubRegion siteSubRegion = SiteSubRegion.Create(this.document, curveLoops, this.TopographySurfaceId);
                                    this.SiteSubRegions.Add(siteSubRegion);
                                }
                                catch (Exception)
                                {
                                    visualizePolygons = true;
                                    this.FailedAttempts++;
                                }
                                createSubRegions.Commit();
                                if (visualizePolygons)
                                {
                                    foreach (Polygon polygon in initialPolygons)
                                    {
                                        this.Visualize(this.document, polygon, 0.0, this.Exponent);
                                    }
                                }
                                if (this.DocumentFileAddress != "")
                                {
                                    this.document.SaveAs(this.DocumentFileAddress, this.opt);
                                }
                            }
                        }
                    }
                }
            }
            #endregion
            
        }
        /// <summary>
        /// Assignes a color-based material to all of the site-subregions
        /// </summary>
        /// <param name="color">Color to be assigned to the site sub-regions</param>
        public void AssignColor(Color color)
        {
            Transaction assignMaterial = new Transaction(this.document, "Assign Materials to SiteSubRegions");
            assignMaterial.Start();
            ElementId matId = Material.Create(this.document, "Subregion");
            Material mat = this.document.GetElement(matId) as Material;
            //Create a new property set that can be used by this material
            StructuralAsset strucAsset = new StructuralAsset("My Property Set", StructuralAssetClass.Concrete);
            strucAsset.Behavior = StructuralBehavior.Isotropic;
            strucAsset.Density = 232.0;
            //Assign the property set to the material.
            PropertySetElement pse = PropertySetElement.Create(this.document, strucAsset);
            mat.SetMaterialAspectByPropertySet(MaterialAspect.Structural, pse.Id);
            mat.Color = color;
            for (int i = 0; i < this.SiteSubRegions.Count; i++)
            {
                foreach (Parameter prm in this.SiteSubRegions[i].TopographySurface.ParametersMap)
                {
                    if (prm.StorageType == StorageType.ElementId && prm.Definition.ParameterType == ParameterType.Material)
                    {
                        prm.Set(matId);
                    }
                }

            }
            assignMaterial.Commit();
        }

        
        #region Polygon boolean operation
        // convert a BoundingBoxUV to a Polygon 
        private Polygon BoundingBoxToPolygon(BoundingBoxUV boundingBox)
        {
            List<XYZ> pnts = new List<XYZ>();
            XYZ p1 = new XYZ(boundingBox.Min.U, boundingBox.Min.V, 0);
            XYZ p2 = new XYZ(boundingBox.Max.U, boundingBox.Min.V, 0);
            XYZ p3 = new XYZ(boundingBox.Max.U, boundingBox.Max.V, 0);
            XYZ p4 = new XYZ(boundingBox.Min.U, boundingBox.Max.V, 0);
            pnts.Add(p1);
            pnts.Add(p2);
            pnts.Add(p3);
            pnts.Add(p4);
            return this.XYZList2Polygon(pnts, this.Exponent);
        }

        //convert a list of XYZ to intPoints
        private Polygon XYZList2Polygon(List<XYZ> lstXYZs, int exponent)
        {
            Polygon contour = new Polygon();
            for (int i = 0; i < lstXYZs.Count; i++)
            {
                contour.Add(XYZ2IntPoint(lstXYZs[i], exponent));
            }
            return contour;
        }
        //convert an XYZ point to an Intpoint
        private IntPoint XYZ2IntPoint(XYZ xyz, int exponent)
        {
            IntPoint pnt = new IntPoint();
            try
            {
                long x = Convert.ToInt64(Math.Floor(xyz.X * Math.Pow(10, exponent)));
                long y = Convert.ToInt64(Math.Floor(xyz.Y * Math.Pow(10, exponent)));
                pnt = new IntPoint(x, y);
            }
            catch (Exception e)
            {
                throw new ArgumentException(e.Message);
            }

            return pnt;
        }

        // polygons treatment
        private Polygons getPolygonsInRegion(int i, int j, double delta)
        {
            //converting the boundingbox of the region to a polygon
            Polygon clip = this.BoundingBoxToPolygon(this.ReagionsBoundingBoxes[i, j]);
            //initial polygons denote part of the subregions that are located within the region
            Polygons initialPolygons = new Polygons();
            Clipper c = new Clipper();
            //c.StrictlySimple = true; // jeremy
            c.AddPolygons(this.IntRoads, PolyType.ptSubject);
            c.AddPolygon(clip, PolyType.ptClip);
            c.Execute(ClipType.ctIntersection, initialPolygons, PolyFillType.pftEvenOdd, PolyFillType.pftEvenOdd);
            // negative Clip is the roads inside the region clipping the region
            Polygons negativeClip = new Polygons();
            Clipper negative = new Clipper();
            //negative.StrictlySimple = true; // jeremy
            negative.AddPolygons(initialPolygons, PolyType.ptClip);
            negative.AddPolygon(clip, PolyType.ptSubject);
            negative.Execute(ClipType.ctIntersection, negativeClip, PolyFillType.pftEvenOdd, PolyFillType.pftEvenOdd);
            //shrinking the negative clip to remove common edges
            Polygons negativeShrank = Clipper.OffsetPolygons(negativeClip, -1 * Math.Pow(10, this.Exponent - 1), ClipperLib.JoinType.jtMiter);
            // expanding the negative shrink the negative clip to remove common edges
            Polygons negativeShrankExpand = Clipper.OffsetPolygons(negativeShrank, Math.Pow(10, this.Exponent - 1) + 3, ClipperLib.JoinType.jtMiter);
            Polygons negativeShrankExpandNegative = new Polygons();
            Clipper negativeAgain = new Clipper();
            //negativeAgain.StrictlySimple = true; // jeremy
            negativeAgain.AddPolygons(negativeShrankExpand, PolyType.ptSubject);
            negativeAgain.AddPolygon(clip, PolyType.ptClip);
            negativeAgain.Execute(ClipType.ctIntersection, negativeShrankExpandNegative, PolyFillType.pftEvenOdd, PolyFillType.pftEvenOdd);

            return negativeShrankExpandNegative;

        }
        //Convert an IntPoint to an XYZ point
        private XYZ IntPoint2XYZ(IntPoint pnt, double z, int exponent)
        {
            double x = Convert.ToDouble(pnt.X) / Math.Pow(10, exponent);
            double y = Convert.ToDouble(pnt.Y) / Math.Pow(10, exponent);
            XYZ xyz = new XYZ(x, y, z);
            return xyz;
        }
        //Convert a list of IntPoints to a list of XYZ points
        private List<XYZ> Polygon2XYZList(Polygon polygon, double z, int exponent)
        {
            List<XYZ> xyzList = new List<XYZ>();
            for (int i = 0; i < polygon.Count; i++)
            {
                xyzList.Add(IntPoint2XYZ(polygon[i], z, exponent));
            }
            return xyzList;
        }
        // Union Polylines
        private Polygons XOR(List<Polygon> polygons)
        {
            Polygons results = new Polygons();
            Clipper c = new Clipper();
            c.AddPolygons(polygons, PolyType.ptSubject);
            //c.StrictlySimple = true;
            c.Execute(ClipType.ctXor, results, PolyFillType.pftEvenOdd, PolyFillType.pftEvenOdd);
            return Clipper.SimplifyPolygons(results);
        }
        public List<ElementId> Visualize(Document doc, Polygon polygon, double elevation, int exponent)
        {
            List<XYZ> points = new List<XYZ>();
            points = this.Polygon2XYZList(polygon, elevation, exponent);
            ProcessPolygon processedContour = new ProcessPolygon(points, true);
            processedContour.LoadEdgeLengths();
            //MessageBox.Show(string.Format("Minimum distance is {0} \n Maximum distance is {1}", processedContour.MinimumEdgeLength.ToString(), processedContour.MaximumEdgeLength.ToString()));
            return processedContour.Visualize(doc);
        }
        /// <summary>
        /// Convert Polygons to a list of curveloops
        /// </summary>
        public CurveLoops PolygonsToCurveLoops(Polygons polygons)
        {
            CurveLoops curveLoops = new CurveLoops(); 
            foreach (Polygon polygon in polygons)
            {
                ProcessPolygon processedPolygon = new ProcessPolygon(this.Polygon2XYZList(polygon, 0.0, this.Exponent),true);
                processedPolygon.RemoveIdenticalPoints();
                processedPolygon.RemoveClosePoints(.5);
                processedPolygon.ForceToFixList();
                processedPolygon.RemoveCollinearity(4 * Math.PI / 180);
                if (processedPolygon.ProcessedPolygon.Count>2)
                {
                    curveLoops.Add(processedPolygon.Get_CurveLoop());
                }
            }
            return curveLoops;
        }

        #endregion

        #region setting ray-tracers to check if a rectangular zone has valid projection on the topography
        // We need to create a 3dview because the when we open a new project file a 3d view does not exist and raytracers need 3dviews
        private View3D createView3d()
        {
            FilteredElementCollector collector0 = new FilteredElementCollector(this.document).OfClass(typeof(View3D));
            foreach (View3D item in collector0)
            {
                if (!item.IsTemplate)
                {
                    return item;
                }
            }
            FilteredElementCollector collector1 = new FilteredElementCollector(this.document);
            collector1 = collector1.OfClass(typeof(ViewFamilyType));
            IEnumerable<ViewFamilyType> viewFamilyTypes = from elem in collector1 let vftype = elem as ViewFamilyType where (vftype.ViewFamily == ViewFamily.ThreeDimensional) select vftype;
            Transaction createView3D = new Transaction(this.document, "CreateView3D");
            createView3D.Start();
            View3D view3D = View3D.CreateIsometric(this.document, viewFamilyTypes.First<ViewFamilyType>().Id);
            if (null != view3D)
            {
                XYZ eye = new XYZ(10, 10, 10);
                XYZ up = new XYZ(0, 1, 1);
                XYZ forward = new XYZ(0, 1, -1);
                ViewOrientation3D viewOrientation3D = new ViewOrientation3D(eye, up, forward);
                view3D.SetOrientation(viewOrientation3D);
                view3D.Name = "RayTracer View";
            }
            createView3D.Commit();
            return view3D;
        }
        //this method checks to see if a rectangular region has valid projection on the ground.
        private bool CanSubregionBeValid(BoundingBoxUV boundingBox)
        {
            List<XYZ> corners = new List<XYZ>();
            XYZ Min = new XYZ(boundingBox.Min.U, boundingBox.Min.V, 0);
            XYZ Max = new XYZ(boundingBox.Max.U, boundingBox.Max.V, 0);
            XYZ p0 = new XYZ(boundingBox.Min.U, boundingBox.Max.V, 0);
            XYZ p1 = new XYZ(boundingBox.Max.U, boundingBox.Min.V, 0);
            corners.Add(Min);
            corners.Add(Max);
            corners.Add(p0);
            corners.Add(p1);
            ReferenceIntersector refIntersector = new ReferenceIntersector(this.TopographySurfaceId, FindReferenceTarget.Mesh, this.RayTracerView);
            foreach (XYZ item in corners)
            {
                ReferenceWithContext referenceWithContext = refIntersector.FindNearest(item, XYZ.BasisZ);
                if (referenceWithContext == null)
                {
                    return false;
                }
            }
            return true;
        }
        #endregion
    
    }

     

}
