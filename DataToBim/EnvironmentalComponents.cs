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

namespace DataToBim
{
    public class EnvironmentalComponents
    {
        /// <summary>
        /// reading building text file
        /// </summary>
        /// <param name="FileAddress"><The address of the file that includes the information of buildings>
        /// <returns></A list of the Building>
        public static List<Building> LoadBuildings(string FileAddress)
        {
            List<Building> buildingList = new List<Building>();
            //reading building text file
            string[] buildingText = File.ReadAllLines(FileAddress);
            for (int i = 0; i < buildingText.Length; i += 4)
            {
                bool insideCampus = true;
                double height = double.Parse(buildingText[i + 1]);
                double area = double.Parse(buildingText[i + 2]);
                Building newBuilding = new Building(height, area);
                //add vertices to the new building
                string[] verticesCoord = buildingText[i + 3].Split(',');
                for (int j = 0; j < verticesCoord.Length; j += 2)
                {
                    if (verticesCoord[j] == "") continue;
                    double X = double.Parse(verticesCoord[j]);
                    double Y = double.Parse(verticesCoord[j + 1]);
                    //skip the buildings that are outside A&M campus
                    if (X > 3558000 || Y < 10204000)
                    {
                        insideCampus = false;
                        break;
                    }
                    XYZ vertex = new XYZ(X, Y, 0);
                    newBuilding.AddVertex(vertex);
                }
                if (insideCampus)
                    buildingList.Add(newBuilding);
            }
            return buildingList;
        }
        public static List<Road> LoadRoads(string FileAddress)
        {
            List<Road> roadList = new List<Road>();
            string[] roadText = File.ReadAllLines(FileAddress);
            for (int i = 0; i < roadText.Length; i += 2)
            {
                Road newRoad = new Road();
                string[] verticesCoord = roadText[i + 1].Split(',');
                for (int j = 0; j < verticesCoord.Length; j += 2)
                {
                    if (verticesCoord[j] == "") continue;
                    double X = double.Parse(verticesCoord[j]);
                    double Y = double.Parse(verticesCoord[j + 1]);
                    XYZ vertex = new XYZ(X, Y, 0);
                    newRoad.AddVertex(vertex);
                }
                roadList.Add(newRoad);
            }
            return roadList;
        }
        public static List<Contour> LoadContours(string FileAddress)
        {
            List<Contour> contourList = new List<Contour>();
            //reading contour text file
            string[] contourText = File.ReadAllLines(FileAddress);
            for (int i = 0; i < contourText.Length; i += 2)
            {
                //use every the other contour
                if (i % 4 == 2) continue;

                Contour newContourline = new Contour();
                string[] verticesCoord = contourText[i + 1].Split(',');
                for (int j = 0; j < verticesCoord.Length; j += 3)
                {
                    if (verticesCoord[j] == "") continue;
                    double X = double.Parse(verticesCoord[j]);
                    double Y = double.Parse(verticesCoord[j + 1]);
                    double Z = double.Parse(verticesCoord[j + 2]);
                    XYZ vertex = new XYZ(X, Y, Z);
                    newContourline.AddVertex(vertex);
                }
                contourList.Add(newContourline);
            }
            return contourList;
        }
    }
    /// <summary>
    /// Building information
    /// </summary>
    public class Building
    {
        public double height { get; set; }
        public double area { get; set; }
        public List<XYZ> vertices = new List<XYZ>();

        public Building(double buildingHeight, double buildingArea)
        {
            this.height = buildingHeight;
            this.area = buildingArea;
        }

        public void AddVertex(XYZ vertex)
        {
            this.vertices.Add(vertex);
        }

    }
    /// <summary>
    /// Contour information
    /// </summary>
    public class Contour
    {
        public List<XYZ> vertices = new List<XYZ>();

        public Contour()
        {
        }

        public void AddVertex(XYZ vertex)
        {
            this.vertices.Add(vertex);
        }
    }
    /// <summary>
    /// Road information
    /// </summary>
    public class Road
    {
        public List<XYZ> vertices = new List<XYZ>();

        public Road()
        {
        }

        public void AddVertex(XYZ vertex)
        {
            this.vertices.Add(vertex);
        }
    }

}
