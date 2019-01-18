using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;

using Security_Check;

namespace PHE_Supports
{
    //Transaction assigned as automatic
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Automatic)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]

    //Creating an external command to provide supports
    public class PHESupports : IExternalCommand
    {
        //instances to store application and the document
        UIDocument m_document = null;

        //boundingbox to store the boundingbox of sectionbox for 3D view
        BoundingBoxXYZ bounds = new BoundingBoxXYZ();

        //3D view to store an active 3D view for intersector to work on
        View3D view3D = null;

        //for storing the selected elements
        ElementSet eleSet = null;

        //tostore the family name in config file, corresopnding family and its family symbol
        string FamilyName = null;
        FamilySymbol symbol = null;
        Family family = null;

        //to check for the security
        bool security = false;

        //for storing the rod length computed, offest fro ends and required spacing
        double rodLength = 0.0, spacing = 0.0;

        //level property of the pipe selected
        Level pipeLevel = null;

        //variable to store element created
        Element tempEle = null;

        //to store the pipe curve of the pipe under selection
        Curve pipeCurve = null;

        //lists to store, spacing details, points for placing instances and list of created instances
        List<XYZ> points = new List<XYZ>();
        List<string> specsData = new List<string>();

        //failure lists
        List<ElementId> FailedToPlace = new List<ElementId>();
        List<ElementId> createdElements = new List<ElementId>();

        //default execute method required by the IExternalCommand class
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                //call to the security check method to check for authentication
                security = SecurityLNT.Security_Check();
                if (security == false)
                {
                    return Result.Succeeded;
                }

                //read the config file
                SpecificationData configFileData = UtilityMethods.ReadConfig();

                //check for discipline
                if (configFileData.discipline != "PHE")
                {
                    TaskDialog.Show("Failed!", "Sorry! Plug-in Not intended for your discipline.");
                    return Result.Succeeded;
                }

                //exception handled
                if (configFileData.offset == -1.0)
                {
                    MessageBox.Show("Configuration data not found!");
                    return Result.Succeeded;
                }

                //get all data from the specification file
                specsData = UtilityMethods.GetAllSpecs(configFileData);

                //exception handled
                if (specsData == null)
                {
                    MessageBox.Show("Specifications not found!");
                    return Result.Succeeded;
                }

                //open  the active document in revit
                m_document = commandData.Application.ActiveUIDocument;

                //get the selected element set
                eleSet = m_document.Selection.Elements;

                if (eleSet.IsEmpty)
                {
                    MessageBox.Show("Please select pipes before executing the Add-in!");
                    return Result.Succeeded;
                }

                //get family name from config data
                FamilyName = configFileData.selectedFamily;

                //check if family exists
                family = UtilityMethods.FindElementByName(m_document.Document, typeof(Family), FamilyName) as Family;

                //if existing
                if (family == null)
                {
                    MessageBox.Show("Please load the family into the project and re-run the Add-in.");
                    return Result.Succeeded;
                }

                //get the 3D view to workon
                view3D = UtilityMethods.Get3D_View(m_document.Document);

                //exception handled
                if (view3D == null)
                {
                    MessageBox.Show("No 3D view available!");
                    return Result.Succeeded;
                }

                //set the 3D view sectionBox to false and copy its bounding box
                bounds = UtilityMethods.GetBounds(view3D);

                //exception handled
                if (bounds == null)
                {
                    MessageBox.Show("Failed to generate bounding box!");
                    return Result.Succeeded;
                }

                //get the family symbol
                symbol = UtilityMethods.GetFamilySymbol(family);

                //exception handled
                if (family == null)
                {
                    MessageBox.Show("No family symbol for the family you specified!");
                    return Result.Succeeded;
                }

                //iterate through all the selected elements
                foreach (Element ele in eleSet)
                {
                    //check whether the selected element is of type pipe
                    if (ele is Pipe)
                    {
                        //get the location curve of the pipe
                        pipeCurve = ((LocationCurve)ele.Location).Curve;

                        //if the length of pipe curve obtained is zero skip that pipe
                        if (pipeCurve.Length == 0)
                        {
                            FailedToPlace.Add(ele.Id);
                            continue;
                        }

                        //from the specification file, get the spacing corresponding to the pipe diameter
                        spacing = UtilityMethods.GetSpacing(ele, specsData);

                        //check if the spacing returned is -1
                        if (spacing == -1)
                        {
                            FailedToPlace.Add(ele.Id);
                            continue;
                        }

                        //get the points for placing the family instances
                        points = UtilityMethods.GetPlacementPoints(spacing, pipeCurve,
                            1000 * configFileData.offset, 1000 * configFileData.minSpacing);

                        //check if the points is null exception
                        if (points == null)
                        {
                            FailedToPlace.Add(ele.Id);
                            continue;
                        }

                        //get the pipe level
                        pipeLevel = (Level)m_document.Document.GetElement(ele.LevelId);

                        //iterate through all the points for placing the family instances
                        foreach (XYZ point in points)
                        {
                            try
                            {
                                //create the instances at each points
                                tempEle = m_document.Document.Create.NewFamilyInstance
                                    (point, symbol, ele, pipeLevel, StructuralType.NonStructural);
                                createdElements.Add(tempEle.Id);
                            }
                            catch
                            {
                                FailedToPlace.Add(ele.Id);
                                continue;
                            }

                            if (configFileData.supportType == "CEILING SUPPORT")
                            {
                                XYZ vector = XYZ.BasisZ;
                                //find the rod length required by using the reference intersector class
                                rodLength = UtilityMethods.ExtendRod(ele, point, m_document.Document, view3D, vector);
                            }

                            if (rodLength == -1)
                            {
                                FailedToPlace.Add(ele.Id);
                                createdElements.Remove(tempEle.Id);
                                m_document.Document.Delete(tempEle.Id);
                                continue;

                            }

                            //adjust the newly created element properties based on the rodlength, 
                            //orientation and dia of pipe
                            if (!UtilityMethods.AdjustElement(m_document.Document,
                                tempEle, point, (Pipe)ele, rodLength, pipeCurve))
                            {
                                FailedToPlace.Add(ele.Id);
                                createdElements.Remove(tempEle.Id);
                                m_document.Document.Delete(tempEle.Id);
                                continue;
                            }
                        }
                    }
                }

                if (!UtilityMethods.SetBounds(bounds, view3D))
                {
                    MessageBox.Show("Sorry! Couldn't restore the sectionbox for the 3D view.");
                    return Result.Succeeded;
                }

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                UtilityMethods.SetBounds(bounds, view3D);
                message = e.Message;
                return Autodesk.Revit.UI.Result.Failed;
            }
            throw new NotImplementedException();
        }
    }
}
