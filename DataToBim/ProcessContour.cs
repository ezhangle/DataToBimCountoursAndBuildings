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
    class ProcessPolygon
    {
        #region Storing initial polygon and processed polygon
        private List<XYZ> initialPolygon;
        /// <summary>
        /// Initial list of XYZ points
        /// </summary>
        public List<XYZ> InitialPolygon { get { return initialPolygon; }}
        /// <summary>
        /// This method reverses all of the changes that were done in the polygon and returns it back to its original list of points
        /// </summary>
        public void Reset()
        {
            if (this.IsClosedPolygon)
            {
                this.processedPolygon = this.initialPolygon.GetRange(0, this.initialPolygon.Count - 1);
            }
            else
            {
                this.processedPolygon = this.initialPolygon;
            }
            this.isPurgedBasedOnCollinearity = false;
            this.isPurgedBasedOnDistance = false;
            this.isInterpolated = false;
            this.edgeLengthsLoaded = false;
        }
        private List<XYZ> processedPolygon = new List<XYZ>();
        /// <summary>
        /// The polygon (i.e. list of points or contours) that was subject to the processings operations
        /// </summary>
        public List<XYZ> ProcessedPolygon {get { return processedPolygon; }}
        public readonly bool IsClosedPolygon;
        #endregion

        #region Interpolation of points on a polygon
        private bool isInterpolated = false;
        public bool IsInterpolated { get { return isInterpolated; } }
        private double interpolationLength;
        public double InterpolationLength {get { return interpolationLength; } }
        //Interpolate XYZ points in a list of XYZ points
        public void Interpolate(double interpolationLength=1.0)
        {
            this.edgeLengthsLoaded = false;
            this.isPurgedBasedOnCollinearity = false;
            this.isPurgedBasedOnDistance = false;
            this.isInterpolated = true;
            this.interpolationLength = interpolationLength;
            List<XYZ> newList = new List<XYZ>();
            for (int i = 0; i < this.processedPolygon.Count - 1; i++)
            {
                newList.AddRange(Inter2XYZ(this.processedPolygon[i], this.processedPolygon[i + 1], interpolationLength));
            }
            if (this.IsClosedPolygon)
            {
                newList.AddRange(Inter2XYZ(this.processedPolygon[this.processedPolygon.Count - 1], this.processedPolygon[0], interpolationLength));
            }
            this.processedPolygon = newList;
        }
        //Interpolate XYZ points between two of XYZ points
        private List<XYZ> Inter2XYZ(XYZ a, XYZ b, double k)
        {
            List<XYZ> MidXYZ = new List<XYZ>();
            double x = a.DistanceTo(b);
            if (x > k)
            {
                int n = Convert.ToInt32(Math.Floor(x / k));
                for (int i = 0; i < n; i++)
                {
                    MidXYZ.Add(a + i * (b - a) / n);
                }
            }
            else
            {
                MidXYZ.Add(a);
            }
            return MidXYZ;
        }
        #endregion
        
        #region Purification of polygon based on a tolerance distance
        private double proximityTolerance;
        /// <summary>
        /// If the polygon was purged this number shows the value of the tolerance assigned to purge the polygon.
        /// </summary>
        public double ProximityTolerance {get { return proximityTolerance; }}
        private bool isPurgedBasedOnDistance = false;
        /// <summary>
        /// Determines whether the polygon has been purged based on the lengths of the edges or not
        /// </summary>
        public bool IsPurgedBasedOnDistance {get { return isPurgedBasedOnDistance; }}
        /// <summary>
        /// Points that are too close will be removed from the list
        /// </summary>
        /// <param name="tolerance">One of the points that is closer than this tolerance will be removed.</param>
        public void RemoveClosePoints(double tolerance)
        {
            this.edgeLengthsLoaded = false;
            this.isInterpolated = false;
            this.isPurgedBasedOnDistance = true;
            this.proximityTolerance = tolerance;
            List<XYZ> newList = new List<XYZ>();
            int newIndex = 0;
            while (newIndex < this.processedPolygon.Count)
            {
                newIndex = nextXYZIndex(processedPolygon, newIndex);
                if (newIndex < processedPolygon.Count)
                {
                    newList.Add(processedPolygon[newIndex]);
                }
            }
            if (newList.Count == 0)
            {
                this.processedPolygon= newList;
            }
            if (newList[newList.Count - 1].DistanceTo(this.processedPolygon[0]) > this.proximityTolerance)
            {
                newList.Insert(0, this.processedPolygon[0]);
            }
            this.processedPolygon = newList;
        }
        //finding the next index for the purifyList function
        private int nextXYZIndex(List<XYZ> list, int index)
        {
            int i = 1;
            while (index + i < list.Count)
            {
                double d = list[index].DistanceTo(list[index + i]);
                if (d < this.proximityTolerance)
                {
                    i++;
                }
                else
                {
                    break;
                }
            }
            return index + i;
        }
        #endregion

        #region Purification of collinearity of edges in a polygon
        private bool isPurgedBasedOnCollinearity = false;
        /// <summary>
        /// Determines whether the polygon has been purged based on the collinearity of its edges or not
        /// </summary>
        public bool IsPurgedBasedOnCollinearity { get { return isPurgedBasedOnCollinearity; }}
        //this is an overload for removing collinearity
        public void RemoveCollinearity(double angle =0.0)
        {
            this.edgeLengthsLoaded = false;
            this.isInterpolated = false;
            this.isPurgedBasedOnCollinearity = true;
            bool treated = this.receivedCollinearityTreatement(angle);
            while (treated && this.processedPolygon.Count>2)
            {
                treated = this.receivedCollinearityTreatement(angle);
            }
        }
        //this method removes collinearity for one point and updates the processedPolygon
        private bool receivedCollinearityTreatement(double angle)
        {
            int i;
            for (i = 0; i < this.processedPolygon.Count; i++)
            {
                if (hasCollinearity(i, angle))
                {
                    this.processedPolygon.RemoveAt(i);
                    this.RemoveClosePoints(.2);
                    return true;
                }
            }
            return false;
        }
        // test to see if one Point has collinearity
        private bool hasCollinearity(int index, double angle)
        {
            if (index == 0 || index == this.processedPolygon.Count-1)
            {
                return false;
            }
            int after = (index == this.processedPolygon.Count-1)? 0: index +1;
            int before = (index == 0) ? this.processedPolygon.Count - 1 : index - 1;
            XYZ bfr = this.processedPolygon[before];
            XYZ pnt = this.processedPolygon[index];
            XYZ aftr = this.processedPolygon[after];
            XYZ direction1 = (pnt - bfr);
            XYZ direction2 = (aftr - pnt);
            if (direction1 == XYZ.Zero || direction2 == XYZ.Zero)
            {
                return true;
            }
            direction1 = direction1.Normalize();
            direction2 = direction2.Normalize();
            double cosine = Math.Abs(direction1.DotProduct(direction2));
            double cosineAngle = Math.Cos(angle);
            bool trueOrFalse = (cosine>=cosineAngle);
            return trueOrFalse;
        }

        #endregion

        #region Maximun and minimum lengths of edges in a polygon
        private bool edgeLengthsLoaded = false;
        /// <summary>
        /// Determines if the values for maximum and minimum lengths are valid
        /// </summary>
        public bool EdgeLengthsLoaded {get { return edgeLengthsLoaded; }}
        private double minimumDistance;
        /// <summary>
        /// The minimum length of polygon edges: use LoadEdgeLengths method before inquiring the minimum and maximum edge lengths
        /// </summary>
        public double MinimumEdgeLength {get { return minimumDistance; }}
        private double maximumDistance;
        /// <summary>
        /// The maximum length of polygon edges: use LoadEdgeLengths method before inquiring the minimum and maximum edge lengths
        /// </summary>
        public double MaximumEdgeLength { get { return maximumDistance; } }
        //Load maximum and minimum edges
        public void LoadEdgeLengths()
        {
            double min = this.processedPolygon[0].DistanceTo(this.processedPolygon[1]);
            double max = min;
            for (int i = 1; i < this.processedPolygon.Count-1; i++)
            {
                double t = this.processedPolygon[i].DistanceTo(this.processedPolygon[i+1]);
                min = (min > t) ? t : min;
                max = (max < t) ? t : max;
            }
            if (this.IsClosedPolygon)
            {
                double t = this.processedPolygon[0].DistanceTo(this.processedPolygon[this.processedPolygon.Count-1]);
                min = (min > t) ? t : min;
                max = (max < t) ? t : max;
            }
            this.maximumDistance = max;
            this.minimumDistance = min;
            this.edgeLengthsLoaded = true;
        }
        #endregion

        #region Flatten
        /// <summary>
        /// Determines if a polygon is planar or not
        /// </summary>
        public bool IsPlanar()
        {
            double height = this.processedPolygon[0].Z;
            for (int i = 1; i < this.processedPolygon.Count; i++)
            {
                if (this.processedPolygon[i].Z != height)
                {
                    return false;
                }
            }
            return true;
        }
        /// <summary>
        /// Flattens a polygon by projecting its points on a plane defined by a z value
        /// </summary>
        /// <param name="elevation">Desired plane defined by a z value</param>
        public void Flatten(double elevation = 0.0)
        {
            this.edgeLengthsLoaded = false;
            List<XYZ> profile = new List<XYZ>();
            for (int i = 0; i < this.processedPolygon.Count; i++)
            {
                profile.Add(new XYZ(this.processedPolygon[i].X, this.processedPolygon[i].Y, elevation));
            }
            this.processedPolygon = profile;
        }

        #endregion

        /// <summary>
        /// Gets a CurveArray representing the polygon
        /// </summary>
        /// <param name="height">A desired plane that is determined by a Z value</param>
        /// <returns>Returns a CurveArray that might be null especifically the polygon is not closed</returns>
        public CurveLoop Get_CurveLoop(double height = 0.0)
        {
            CurveLoop crvAr = new CurveLoop();
            for (int i = 0; i < this.processedPolygon.Count; i++)
            {
                int j = (i == this.processedPolygon.Count - 1) ? 0 : i + 1;
                try
                {
                    Line l = Line.CreateBound(this.processedPolygon[i], this.processedPolygon[j]);
                    crvAr.Append(l);
                }
                catch (Exception)
                {
                    return null;
                }
            }
            return crvAr;
        }

        /// <summary>
        /// Gets a ReferenceArray representing the polygon
        /// </summary>
        /// <param name="document">The document in which the ReferenceArray is going to be generated</param>
        /// <param name="height">A desired plane that is determined by a Z value</param>
        /// <returns>Returns a ReferenceArray that might be null especifically the polygon is not closed</returns>
        public ReferenceArray Get_ReferenceArray(Document document, XYZ translation)
        {
            Transaction createReferenceArray = new Transaction(document, "Create ReferenceArray");
            FailureHandlingOptions failOpt = createReferenceArray.GetFailureHandlingOptions();
            failOpt.SetFailuresPreprocessor(new WarningSwallower());
            createReferenceArray.SetFailureHandlingOptions(failOpt);
            createReferenceArray.Start();
            ReferenceArray refAr = new ReferenceArray();
            Plane p = new Plane(XYZ.BasisZ, new XYZ(this.processedPolygon[0].X, this.processedPolygon[0].Y, 0));
            SketchPlane skp = SketchPlane.Create(document, p);
            try
            {
                for (int i = 0; i < this.processedPolygon.Count; i++)
                {
                    int j = (i == this.processedPolygon.Count - 1) ? 0 : i + 1;
                    XYZ ptA = this.processedPolygon[i] - translation;
                    XYZ PtB = this.processedPolygon[j] - translation;
                    Line l = Line.CreateBound(ptA, PtB);
                    ModelCurve mCrv = document.FamilyCreate.NewModelCurve(l, skp);
                    refAr.Append(mCrv.GeometryCurve.Reference);
                }
            }
            catch (Exception)
            {
                createReferenceArray.Commit();
                return null;
            }
            createReferenceArray.Commit();
            return refAr;
        }
        /// <summary>
        /// This constroctor will decide if the polygon is closed or not
        /// </summary>
        /// <param name="points">List of XYZ points</param>
        public ProcessPolygon(List<XYZ> points)
        {
            this.initialPolygon = points;
            this.IsClosedPolygon = points[0].IsAlmostEqualTo(points[points.Count - 1]);
            if (this.IsClosedPolygon)
            {
                this.processedPolygon = points.GetRange(0, points.Count - 1);
            }
            else
            {
                this.processedPolygon = points;
            }
        }
        /// <summary>
        /// This constroctor asks if the polygon is closed or not
        /// </summary>
        /// <param name="points">List of XYZ points</param>
        /// <param name="closed">Whether the polygon is closed or not</param>
        public ProcessPolygon(List<XYZ> points, bool closed)
        {
            this.initialPolygon = points;
            this.IsClosedPolygon = closed;
            this.processedPolygon = points;
        }
        /// <summary>
        /// Draws model curves to represent the polygon
        /// </summary>
        /// <param name="doc">Name of the document to draw the polygon</param>
        /// <param name="elevation">represents the elevation of a plane to draw the polygon</param>
        /// <returns>A list of element IDs referring to the polygon lines</returns>
        public List<ElementId> Visualize(Document doc, double elevation = 0.0)
        {
            List<ElementId> lines = new List<ElementId>();
            Transaction t = new Transaction(doc, "Draw a contour");
            t.Start();
            FailureHandlingOptions failOpt = t.GetFailureHandlingOptions();
            failOpt.SetFailuresPreprocessor(new WarningSwallower());
            t.SetFailureHandlingOptions(failOpt);
            Plane p = doc.Application.Create.NewPlane(XYZ.BasisZ, new XYZ(0, 0, elevation));
            SketchPlane skp = SketchPlane.Create(doc, p);
            for (int i = 0; i < this.processedPolygon.Count - 1; i++)
            {
                try
                {
                    XYZ p1 = new XYZ(this.processedPolygon[i].X, this.processedPolygon[i].Y, elevation);
                    XYZ p2 = new XYZ(this.processedPolygon[i + 1].X, this.processedPolygon[i + 1].Y, elevation);
                    double d = this.processedPolygon[i].DistanceTo(this.processedPolygon[i + 1]);
                    Line line = Line.CreateBound(p1, p2);
                    ModelLine l = doc.Create.NewModelCurve(line, skp) as ModelLine;
                    lines.Add(l.Id);
                }
                catch (Exception) {}
            }

            try
            {
                XYZ p1 = new XYZ(this.processedPolygon[this.processedPolygon.Count - 1].X, this.processedPolygon[this.processedPolygon.Count - 1].Y, elevation);
                XYZ p2 = new XYZ(this.processedPolygon[0].X, this.processedPolygon[0].Y, elevation);
                Line line = Line.CreateBound(p1, p2);
                ModelLine l = doc.Create.NewModelCurve(line, skp) as ModelLine;
                lines.Add(l.Id);
            }
            catch (Exception) {}
            t.Commit();
            return lines;
        }

        #region Force fix
        /// <summary>
        /// this method removes the points from polygon which cause problem for Revit to draw it
        /// </summary>
        private List<XYZ> FixedOnetime(List<XYZ> pnts)
        {
            bool[] exist = new bool[pnts.Count];
            for (int i = 0; i < pnts.Count - 1; i++)
            {
                try
                {
                    Line l = Line.CreateBound(pnts[i], pnts[i + 1]);
                    exist[i] = true;
                }
                catch (Exception)
                {
                    exist[i] = false;
                }
            }
            try
            {
                Line l = Line.CreateBound(pnts[pnts.Count - 1], pnts[0]);
                exist[pnts.Count - 1] = true;
            }
            catch (Exception)
            {
                exist[pnts.Count - 1] = false;
            }
            List<XYZ> newPnts = new List<XYZ>();
            for (int i = 0; i < pnts.Count; i++)
            {
                if (exist[i])
                {
                    newPnts.Add(pnts[i]);
                }
            }
            return newPnts;
        }
        /// <summary>
        /// This function recursively remove some points from a list untill a curveloop can be created for the list or the list includes less than 3 points.
        /// </summary>
        /// <param name="pnts">A list to be fixed</param>
        /// <returns></returns>
        public void ForceToFixList()
        {
            List<XYZ> points = this.processedPolygon;
            while (points.Count > FixedOnetime(points).Count)
            {
                points = FixedOnetime(points);
                if (points.Count < 3)
                {
                    this.processedPolygon = points;
                    return;
                }
            }
            this.processedPolygon = points;
            this.edgeLengthsLoaded = false;
            this.isInterpolated = false;
            this.isPurgedBasedOnCollinearity = false;
            this.isPurgedBasedOnDistance = false;
        }
        #endregion

        #region Identical points treatment
        //this method returns the number of identical points in a polygon
        public int NumberOfIdenticalPoints()
        {
            int counter = 0;
            Dictionary<string, int> dict = new Dictionary<string, int>();
            for (int i = 0; i < this.processedPolygon.Count; i++)
            {
                string key = this.processedPolygon[i].X.ToString() + this.processedPolygon[i].Y.ToString();
                try
                {
                    dict.Add(key, i);
                }
                catch (Exception)
                {
                    counter++;
                }
            }
            return counter;
        }
        /// <summary>
        /// This method removes the identical points on a polygon
        /// </summary>
        public void RemoveIdenticalPoints()
        {
            List<XYZ> pnts = new List<XYZ>();
            Dictionary<string, int> dict = new Dictionary<string, int>();
            for (int i = 0; i < this.processedPolygon.Count; i++)
            {
                string key = this.processedPolygon[i].X.ToString() + this.processedPolygon[i].Y.ToString();
                try
                {
                    dict.Add(key, i);
                    pnts.Add(this.processedPolygon[i]);
                }
                catch (Exception) { }
            }
            this.processedPolygon = pnts;
            this.edgeLengthsLoaded = false;
            this.isInterpolated = false;
            this.isPurgedBasedOnCollinearity = false;
            this.isPurgedBasedOnDistance = false;
        }
        #endregion

        public static BoundingBoxUV GetRange(List<XYZ> points)
        {
            double minX = points[0].X, maxX = points[0].X, minY = points[0].Y, maxY = points[0].Y;
            for (int i = 1; i < points.Count; i++)
            {
                minX = (minX > points[i].X)? points[i].X: minX;
                maxX = (maxX<points[i].X)? points[i].X: maxX;
                minY = (minY>points[i].Y)? points[i].Y: minY;
                maxY = (maxY < points[i].Y) ? points[i].Y : maxY;
            }
            return new BoundingBoxUV(minX, minY, maxX, maxY);
        }

    }

}
