//
// (C) Copyright 2003-2010 by Autodesk, Inc.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE. AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//

 
using System;
using System.Data;
using System.Collections.Generic;
using System.Text;

using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Reflection;

using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Structure;


using System.Drawing.Drawing2D;
using System.Threading;
using Autodesk.Revit.UI.Selection;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
//using AnalyticalSupportData_info;

using Panel = Autodesk.Revit.DB.Panel;
using Element = Autodesk.Revit.DB.Element;
using Instance = Autodesk.Revit.DB.Instance;

//using AnalyticalSupportData_Info;
//using DynamicTable;
using AnalyticalSupportData_Info;
using AnalyticalSupportData_info;
//using AnalyticalSupportData_Info;


namespace Revit.SDK.Samples.AnalyticalSupportData_Info.CS
{
    
    /// <summary>
    /// get element's id and type information and its supported information.
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class Command: IExternalCommand
    {
        //public static ExternalCommandData m_revit = null;
        private static ExternalCommandData m_revit = null;
        //ExternalCommandData m_revit    = null;  // application of Revit
       

        const double ToMetricUnit = 0.3048;
        //const double ToMetricThickness = 0.3048;
        const double ToMetricUnitWeight = 0.010764;            //coefficient of converting unit weight from internal unit to metric unit 
        const double ToMetricStress = 0.334554;            //coefficient of converting stress from internal unit to metric unit
        const double ToImperialUnitWeight = 6.365827;            //coefficient of converting unit weight from internal unit to imperial unit
        const double ChangedUnitWeight = 14.5;                //the value of unit weight of selected component to be set
        //double ToMetricArea = ToMetricThickness * ToMetricThickness;
        //double BuildingSurfaceArea;
        //double BuildingPlanArea;
        //double BuildingHeight;
        //double M_NetVolume;
        //double M_NetArea;
        //string M_QuantityType;

        //string s_Length;



        //public GeomHelper()  //:this()
        //{
        //    m_currentDoc = Command.CommandData.Application.ActiveUIDocument.Document;

        //    //C.slabLength = SlabLength;
        //}

        /// <summary>
        /// property to get private member variable m_elementInformation.
        /// </summary>
       
       


        DataTable m_materialQuantitiesTable = null;

        public DataTable MaterialQuantitiesTable
        {
            get
            {
                return m_materialQuantitiesTable;
            }


        }

        DataTable m_CreatedQuantitiesTable = null;
        public DataTable MCreated_QuantitiesTable
        {
            get
            {
                return m_CreatedQuantitiesTable;
            }


        }

       



        //static AddInId appId = new AddInId(new Guid("7E5CAC0D-F3D8-4040-89D6-0828D681561B"));


        DataTable Created_QuantitiesTable = null;


        /// <summary>
        /// Implement this method as an external command for Revit.
        /// </summary>
        /// <param name="revit">An object that is passed to the external application 
        /// which contains data related to the command, 
        /// such as the application object and active view.</param>
        /// <param name="message">A message that can be set by the external application 
        /// which will be displayed if a failure or cancellation is returned by 
        /// the external command.</param>
        /// <param name="elements">A set of elements to which the external application 
        /// can add elements that are to be highlighted in case of failure or cancellation.</param>
        /// <returns>Return the status of the external command. 
        /// A result of Succeeded means that the API external method functioned as expected. 
        /// Cancelled can be used to signify that the user cancelled the external operation 
        /// at some point. Failure should be returned if the application is unable to proceed with 
        /// the operation.</returns>
        public Autodesk.Revit.UI.Result Execute(Autodesk.Revit.UI.ExternalCommandData revit,
                                                              ref string message,
                                                              ElementSet elements)
        {




            Autodesk.Revit.ApplicationServices.Application app = revit.Application.Application;
            m_doc = revit.Application.ActiveUIDocument.Document;

            //String filename = "CalculateMaterialQuantities.txt";

            //m_writer = new StreamWriter(filename);

            //ExecuteCalculationsWith<RoofMaterialQuantityCalculator>();
            //ExecuteCalculationsWith<WallMaterialQuantityCalculator>();
            //ExecuteCalculationsWith<FloorMaterialQuantityCalculator>();

            //m_writer.Close();
            //DataTable CreatedQuantitiesTable = CreateMaterialDataTable();


          

            String filename = @"C:\CalculateMaterialQuantities_" + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_").Replace(" ", "_") + ".csv";

            m_writer = new StreamWriter(filename);          

            //m_writer.WriteLine(String.Format("{0},{1:F2},{2:F2}",
            //       "Material Description",  // Element names may have ',' in them
            //       "Net Volume (cubic m)", "Net Area (cubic m)"));

            //m_writer.WriteLine(String.Format("Material Description,{1}", GetElementTypeName(), legendLine));

            ExecuteCalculationsWith<RoofMaterialQuantityCalculator>();
            ExecuteCalculationsWith<WallMaterialQuantityCalculator>();
            ExecuteCalculationsWith<DoorMaterialQuantityCalculator>();
            ExecuteCalculationsWith<WindowMaterialQuantityCalculator>();
            ExecuteCalculationsWith<FloorMaterialQuantityCalculator>();
            ExecuteCalculationsWith<FoundationMaterialQuantityCalculator>();
            // ExecuteCalculationsWith<WallFinishMaterialQuantityCalculator>();

            

            m_writer.Close();

            

           // Process.Start(@"EXCEL", filename);

            ////////////////
            ///////////////

            // convert csv to DataTable

            string[] Lines = File.ReadAllLines(filename);
            string[] Fields;
            Fields = Lines[0].Split(new char[] { ',' });
            int Cols = Fields.GetLength(0);

            DataTable dt = new DataTable();

            //1st row must be column names; force lower case to ensure matching later on.
            for (int i = 0; i < Cols; i++)
                dt.Columns.Add(Fields[i].ToLower(), typeof(string));
            DataRow Row;
            for (int i = 1; i < Lines.GetLength(0); i++)
            {
                //double fieldValue = 0.00;

                Fields = Lines[i].Split(new char[] { ',' });
                Row = dt.NewRow();
                for (int f = 0; f < Cols; f++)
                    //{
                    //    //if (Fields[f] != "")
                    //    //{
                    //    //    fieldValue = Double.Parse(Fields[f]);
                    //    //    Row[f] = fieldValue.ToString("N2");
                    //    //}

                    //}

                    Row[f] = Fields[f];

                dt.Rows.Add(Row);
            }

            Created_QuantitiesTable = dt;


            m_materialQuantitiesTable = Created_QuantitiesTable;


            using (Embodied_Energy_and_Carbon f = new Embodied_Energy_and_Carbon())
            {
                f.MaterialQuantitiesTable_Value = MaterialQuantitiesTable;
                // f.ShowDialog();
            }



            // Set currently executable application to private variable m_revit
            m_revit = revit;
            //MessageBox.Show("Length3: " + (s_slabLen * ToMetricUnit).ToString("F2"));

            //ElementSet selectedElements = m_revit.Application.ActiveUIDocument.Selection.Elements;

            ElementSet selectedElements = new ElementSet();

            foreach (ElementId elementId in m_revit.Application.ActiveUIDocument.Selection.GetElementIds())
            {
                selectedElements.Insert(m_revit.Application.ActiveUIDocument.Document.GetElement(elementId));
            }

           // GeomHelper h = new GeomHelper();

            //slabLength = h.SlabLength;

            //MessageBox.Show("Length6: " + (s_slabLen * ToMetricUnit).ToString("F2"));

            // show UI
            OperationMode displayForm = new OperationMode(this);
            {
                // MessageBox.Show("Length3.1: " + (s_slabLen * ToMetricUnit).ToString("F2"));
                displayForm.ShowDialog();
                //MessageBox.Show("Length3.2: " + (s_slabLen * ToMetricUnit).ToString("F2"));


            }

            //////////////
            /////////////




            // return Result.Cancelled;
            return Autodesk.Revit.UI.Result.Succeeded;
        }

        private void ExecuteCalculationsWith<T>() where T : MaterialQuantityCalculator, new()
        {
          
            T calculator = new T();
            calculator.SetDocument(m_doc);
            calculator.CalculateMaterialQuantities();
            calculator.ReportResults(m_writer);
            
           
        }

        #region Basic Command Data
        private Document m_doc;
        private TextWriter m_writer;
       
       
        
        #endregion



        public Autodesk.Revit.DB.DisplayUnitType UnitType;



        /// <summary>
        /// ExternalCommandData
        /// </summary>
        public static ExternalCommandData CommandData
        {
            get
            {
                return m_revit;
            }
        }




        /// <summary>
        /// get all the required information of selected elements and store them in a data table
        /// </summary>
        /// <param name="selectedElements">
        /// all selected elements in Revit main program
        /// </param>
        /// <returns>
        /// a data table which store all the required information
        /// </returns>
       

        


        //double BuildingSurfaceArea;
        private static double deck_Length;
        private static double deck_Width;

        public double Slab_Length
        {
            set
            {
                deck_Length = value;
            }
        }

        public double Slab_Width
        {
            set
            {
                deck_Width = value;
            }
        }


        
        private DataTable CreateMaterialDataTable()
        {
            // Create a new DataTable.
            // DataTable elementInformationTable = new DataTable("ElementInformationTable");

            DataTable QuantitiesTable = new DataTable();

            // Create element unique id column and add to the DataTable.
            DataColumn GifaTypeColumn = new DataColumn();
            GifaTypeColumn.DataType = typeof(System.String);
            GifaTypeColumn.ColumnName = "Gifa Type";
            GifaTypeColumn.Caption = "Gifa Type";
            GifaTypeColumn.ReadOnly = true;
            QuantitiesTable.Columns.Add(GifaTypeColumn);


            // Create element unique id column and add to the DataTable.
            DataColumn QuantityTypeColumn = new DataColumn();
            QuantityTypeColumn.DataType = typeof(System.String);
            QuantityTypeColumn.ColumnName = "Quantity Type";
            QuantityTypeColumn.Caption = "Quantity Type";
            QuantityTypeColumn.ReadOnly = true;
            QuantitiesTable.Columns.Add(QuantityTypeColumn);

            // Create element unique id column and add to the DataTable.
            DataColumn NetVolumeColumn = new DataColumn();
            NetVolumeColumn.DataType = typeof(System.String);
            NetVolumeColumn.ColumnName = "Net Volume";
            NetVolumeColumn.Caption = "Net Volume";
            NetVolumeColumn.ReadOnly = true;
            QuantitiesTable.Columns.Add(NetVolumeColumn);

            return QuantitiesTable;
        }


        public string s { get; set; }




    }

    ///////////////
    ////////////////
    //////////////////
    //////////////////////////
    /////////////////////////////

    
    /// <summary>
    /// The wall material quantity calculator specialized class.
    /// </summary>
    class WallMaterialQuantityCalculator : MaterialQuantityCalculator
    {
        protected override void CollectElements()
        {
            // filter for non-symbols that match the desired category so that inplace elements will also be found
            FilteredElementCollector collector = new FilteredElementCollector(m_doc);
            m_elementsToProcess = collector.OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements();
        }

        protected override string GetElementTypeName()
        {
            return "Wall";
        }
    }

    class DoorMaterialQuantityCalculator : MaterialQuantityCalculator
    {
        protected override void CollectElements()
        {
            // filter for non-symbols that match the desired category so that inplace elements will also be found
            FilteredElementCollector collector = new FilteredElementCollector(m_doc);
            m_elementsToProcess = collector.OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements();
        }

        protected override string GetElementTypeName()
        {
            return "Door";
        }
    }

    class WindowMaterialQuantityCalculator : MaterialQuantityCalculator
    {
        protected override void CollectElements()
        {
            // filter for non-symbols that match the desired category so that inplace elements will also be found
            FilteredElementCollector collector = new FilteredElementCollector(m_doc);
            m_elementsToProcess = collector.OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToElements();
        }

        protected override string GetElementTypeName()
        {
            return "Window";
        }
    }




    /// <summary>
    /// The floor material quantity calculator specialized class.
    /// </summary>
    class FloorMaterialQuantityCalculator : MaterialQuantityCalculator
    {
        protected override void CollectElements()
        {
            FilteredElementCollector collector = new FilteredElementCollector(m_doc);
            m_elementsToProcess = collector.OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements();
        }

        protected override string GetElementTypeName()
        {
            return "Floor";
        }
    }

    /// <summary>
    /// The roof material quantity calculator specialized class.
    /// </summary>
    class RoofMaterialQuantityCalculator : MaterialQuantityCalculator
    {
        protected override void CollectElements()
        {
            FilteredElementCollector collector = new FilteredElementCollector(m_doc);
            m_elementsToProcess = collector.OfCategory(BuiltInCategory.OST_Roofs).WhereElementIsNotElementType().ToElements();
        }

        protected override string GetElementTypeName()
        {
            return "Roof";
        }
    }

    ////////////////


    class FoundationMaterialQuantityCalculator : MaterialQuantityCalculator
    {
        protected override void CollectElements()
        {
            FilteredElementCollector collector = new FilteredElementCollector(m_doc);
            m_elementsToProcess = collector.OfCategory(BuiltInCategory.OST_StructuralFoundation).WhereElementIsNotElementType().ToElements();
        }

        protected override string GetElementTypeName()
        {
            return "Foundation";
        }
    }


    //class WallFinishMaterialQuantityCalculator : MaterialQuantityCalculator
    //{
    //    protected override void CollectElements()
    //    {
    //        FilteredElementCollector collector = new FilteredElementCollector(m_doc);
    //        m_elementsToProcess = collector.OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements();
    //    }

    //    protected override string GetElementTypeName()
    //    {
    //        return "WallFinish";
    //    }
    //}

    //Command: IExternalCommand
    /////////////////////

    /// <summary>
    /// The base material quantity calculator for all element types.
    /// </summary>
    abstract class MaterialQuantityCalculator 
    {
        /// <summary>
        /// The list of elements for material quantity extraction.
        /// </summary>
        protected IList<Element> m_elementsToProcess;

        /// <summary>
        /// Override this to populate the list of elements for material quantity extraction.
        /// </summary>
        protected abstract void CollectElements();

        /// <summary>
        /// Override this to return the name of the element type calculated by this calculator.
        /// </summary>
        protected abstract String GetElementTypeName();

        /// <summary>
        /// Sets the document for the calculator class.
        /// </summary>
        public void SetDocument(Document d)
        {
            m_doc = d;
            Autodesk.Revit.ApplicationServices.Application app = d.Application;
        }

        /// <summary>
        /// Executes the calculation.
        /// </summary>
        public void CalculateMaterialQuantities()
        {
            CollectElements();
            CalculateNetMaterialQuantities();
            CalculateGrossMaterialQuantities();
        }

        /// <summary>
        /// Calculates net material quantities for the target elements.
        /// </summary>
        private void CalculateNetMaterialQuantities()
        {
            foreach (Element e in m_elementsToProcess)
            {
                CalculateMaterialQuantitiesOfElement(e);
            }
        }

        /// <summary>
        /// Calculates gross material quantities for the target elements (material quantities with 
        /// all openings, doors and windows removed). 
        /// </summary>
        private void CalculateGrossMaterialQuantities()
        {
            m_calculatingGrossQuantities = true;
            Transaction t = new Transaction(m_doc);
            t.SetName("Delete Cutting Elements");
            t.Start();
            // DeleteAllCuttingElements();
            m_doc.Regenerate();
            foreach (Element e in m_elementsToProcess)
            {
                CalculateMaterialQuantitiesOfElement(e);
            }
            t.RollBack();
        }

        /// <summary>
        /// Delete all elements that cut out of target elements, to allow for calculation of gross material quantities.
        /// </summary>
        private void DeleteAllCuttingElements()
        {
            IList<ElementFilter> filterList = new List<ElementFilter>();
            FilteredElementCollector collector = new FilteredElementCollector(m_doc);

            // (Type == FamilyInstance && (Category == Door || Category == Window) || Type == Opening
            ElementClassFilter filterFamilyInstance = new ElementClassFilter(typeof(FamilyInstance));
            ElementCategoryFilter filterWindowCategory = new ElementCategoryFilter(BuiltInCategory.OST_Windows);
            ElementCategoryFilter filterDoorCategory = new ElementCategoryFilter(BuiltInCategory.OST_Doors);
            LogicalOrFilter filterDoorOrWindowCategory = new LogicalOrFilter(filterWindowCategory, filterDoorCategory);
            LogicalAndFilter filterDoorWindowInstance = new LogicalAndFilter(filterDoorOrWindowCategory, filterFamilyInstance);

            ElementClassFilter filterOpening = new ElementClassFilter(typeof(Opening));

            LogicalOrFilter filterCuttingElements = new LogicalOrFilter(filterOpening, filterDoorWindowInstance);
            ICollection<Element> cuttingElementsList = collector.WherePasses(filterCuttingElements).ToElements();

            foreach (Element e in cuttingElementsList)
            {
                // Doors in curtain grid systems cannot be deleted.  This doesn't actually affect the calculations because
                // material quantities are not extracted for curtain systems.
                if (e.Category != null)
                {
                    if (e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Doors)
                    {
                        FamilyInstance door = e as FamilyInstance;
                        Wall host = door.Host as Wall;

                        if (host.CurtainGrid != null)
                            continue;
                    }
                    ICollection<ElementId> deletedElements = m_doc.Delete(e.Id);

                    // Log failed deletion attempts to the output.  (These may be other situations where deletion is not possible but 
                    // the failure doesn't really affect the results.
                    if (deletedElements == null || deletedElements.Count < 1)
                    {
                        m_warningsForGrossQuantityCalculations.Add(
                                 String.Format("   The tool was unable to delete the {0} named {2} (id {1})", e.GetType().Name, e.Id, e.Name));
                    }
                }
            }
        }

        /// <summary>
        /// Store calculated material quantities in the storage collection.
        /// </summary>
        /// <param name="materialId">The material id.</param>
        /// <param name="volume">The extracted volume.</param>
        /// <param name="area">The extracted area.</param>
        /// <param name="quantities">The storage collection.</param>
        private void StoreMaterialQuantities(ElementId materialId, double volume, double area,
                                            Dictionary<ElementId, MaterialQuantities> quantities)
        {
            ///////////////


            //const double ToMetricUnit = 0.3048;
            ////const double ToMetricThickness = 0.3048;
            //const double ToMetricUnitWeight = 0.010764;            //coefficient of converting unit weight from internal unit to metric unit 
            //const double ToMetricStress = 0.334554;            //coefficient of converting stress from internal unit to metric unit
            //const double ToImperialUnitWeight = 6.365827;            //coefficient of converting unit weight from internal unit to imperial unit
            //const double ChangedUnitWeight = 14.5;                //the value of unit weight of selected component to be set
            ////double ToMetricArea = ToMetricThickness * ToMetricThickness;


            ////////////////////



            MaterialQuantities materialQuantityPerElement;
            bool found = quantities.TryGetValue(materialId, out materialQuantityPerElement);

            if (found)
            {
                if (m_calculatingGrossQuantities)
                {
                    materialQuantityPerElement.GrossVolume += volume;
                    materialQuantityPerElement.GrossArea += area;
                }
                else
                {
                    materialQuantityPerElement.NetVolume += volume;
                    materialQuantityPerElement.NetArea += area;
                }
            }
            else
            {
                materialQuantityPerElement = new MaterialQuantities();
                if (m_calculatingGrossQuantities)
                {
                    materialQuantityPerElement.GrossVolume = volume;
                    materialQuantityPerElement.GrossArea = area;
                }
                else
                {
                    materialQuantityPerElement.NetVolume = volume;
                    materialQuantityPerElement.NetArea = area;
                }
                quantities.Add(materialId, materialQuantityPerElement);
            }
           
        }

        /// <summary>
        /// Calculate and store material quantities for a given element.
        /// </summary>
        /// <param name="e">The element.</param>
        private void CalculateMaterialQuantitiesOfElement(Element e)
        {
            ElementId elementId = e.Id;
            ICollection<ElementId> materials = e.GetMaterialIds(false);
            const double ToMetricUnit = 0.3048;

            foreach (ElementId materialId in materials)
            {
                double volume = e.GetMaterialVolume(materialId) * ToMetricUnit * ToMetricUnit * ToMetricUnit;
                double area = e.GetMaterialArea(materialId, false) * ToMetricUnit * ToMetricUnit;

                if (volume > 0.0 || area > 0.0)
                {
                    StoreMaterialQuantities(materialId, volume, area, m_totalQuantities);

                    Dictionary<ElementId, MaterialQuantities> quantityPerElement;
                    bool found = m_quantitiesPerElement.TryGetValue(elementId, out quantityPerElement);
                    if (found)
                    {
                        StoreMaterialQuantities(materialId, volume, area, quantityPerElement);
                    }
                    else
                    {
                        quantityPerElement = new Dictionary<ElementId, MaterialQuantities>();
                        StoreMaterialQuantities(materialId, volume, area, quantityPerElement);
                        m_quantitiesPerElement.Add(elementId, quantityPerElement);
                    }
                }
            }
        }

        /// <summary>
        /// Write results in CSV format to the indicated output writer.
        /// </summary>
        /// <param name="writer">The output text writer.</param>
        public void ReportResults(TextWriter writer)
        {
            if (m_totalQuantities.Count == 0)
                return;

            //String legendLine = "Gross volume(cubic ft),Net volume(cubic ft),Gross area(sq ft),Net area(sq ft)";
            //String legendLine = "Net volume (cubic m), Net area (sq m)";

            //writer.WriteLine();
            //writer.WriteLine(String.Format("Totals for {0} elements,{1}", GetElementTypeName(), legendLine));

            // If unexpected deletion failures occurred, log the warnings to the output.
            if (m_warningsForGrossQuantityCalculations.Count > 0)
            {
                writer.WriteLine("WARNING: Calculations for gross volume and area may not be completely accurate due to the following warnings: ");
                foreach (String s in m_warningsForGrossQuantityCalculations)
                    writer.WriteLine(s);
                writer.WriteLine();
                
            }


          

            ReportResultsFor(m_totalQuantities, writer);

           
        }


        

        //DataTable CreatedQuantitiesTable = CreateMaterialDataTable();
        /// <summary>
        /// Write the contents of one storage collection to the indicated output writer.
        /// </summary>
        /// <param name="quantities">The storage collection for material quantities.</param>
        /// <param name="writer">The output writer.</param>
        private void ReportResultsFor(Dictionary<ElementId, MaterialQuantities> quantities, TextWriter writer)
        {


            
            foreach (ElementId keyMaterialId in quantities.Keys)
            {
                

                ElementId materialId = keyMaterialId;
                MaterialQuantities quantity = quantities[materialId];

                Material material = m_doc.GetElement(materialId) as Material;

                //writer.WriteLine(String.Format("   {0} Net: [{1:F2} cubic ft {2:F2} sq. ft]  Gross: [{3:F2} cubic ft {4:F2} sq. ft]", material.Name, quantity.NetVolume, quantity.NetArea, quantity.GrossVolume, quantity.GrossArea));
                writer.WriteLine(String.Format("{0},{1:F2},{2:F2}",
                    GetElementTypeName() + " - " +  material.Name.Replace(',', ':'),  // Element names may have ',' in them
                    quantity.NetVolume, quantity.NetArea));
                               
                

                M_NetVolume = quantity.NetVolume;
                //M_NetArea;
                M_QuantityType = material.Name.Replace(',', ':');

                M_GifaType = GetElementTypeName();


                using (Embodied_Energy_and_Carbon f = new Embodied_Energy_and_Carbon ())
                {
                    
                    f.MaterialNetVolume_Value = MaterialNetVolume;
                    f.MaterialQuantityType_Value = MaterialQuantityType;
                    f.MaterialGifaType_Value = MaterialGifaType;

                    // f.ShowDialog();
                }

               
            }

          




        }


       

        #region Results Storage
        /// <summary>
        /// A storage of material quantities per individual element.
        /// </summary>
        private Dictionary<ElementId, Dictionary<ElementId, MaterialQuantities>> m_quantitiesPerElement = new Dictionary<ElementId, Dictionary<ElementId, MaterialQuantities>>();

        /// <summary>
        /// A storage of material quantities for the entire project.
        /// </summary>
        private Dictionary<ElementId, MaterialQuantities> m_totalQuantities = new Dictionary<ElementId, MaterialQuantities>();
        //public Dictionary<ElementId, MaterialQuantities> m_totalQuantities = new Dictionary<ElementId, MaterialQuantities>();
        /// <summary>
        /// Flag indicating the mode of the calculation.
        /// </summary>
        private bool m_calculatingGrossQuantities = false;

        /// <summary>
        /// A collection of warnings generated due to failure to delete elements in advance of gross quantity calculations.
        /// </summary>
        private List<String> m_warningsForGrossQuantityCalculations = new List<string>();
        #endregion

        protected Document m_doc;



        //DataTable CreatedQuantitiesTable = null ;

        double M_NetVolume;
        //double M_NetArea;
        string M_QuantityType;
        string M_GifaType;


        public double MaterialNetVolume
        {
            get
            {
                return M_NetVolume;
            }
        }
        //public double MateralNetArea
        //{
        //    get
        //    {
        //        return M_NetArea;
        //    }
        //}

        public string MaterialQuantityType
        {
            get
            {
                return M_QuantityType;
            }
        }

        public string MaterialGifaType
        {
            get
            {
                return M_GifaType;
            }
        }


        DataTable m_materialQuantitiesTable = null;

        public DataTable MaterialQuantitiesTable
        {
            get
            {
                return m_materialQuantitiesTable;
            }


        }

        
    }

    /// <summary>
    /// A storage class for the extracted material quantities.
    /// </summary>
    class MaterialQuantities
    {
        /// <summary>
        /// Gross volume (cubic ft)
        /// </summary>
        public double GrossVolume { get; set; }

        /// <summary>
        /// Gross area (sq. ft)
        /// </summary>
        public double GrossArea { get; set; }

        /// <summary>
        /// Net volume (cubic ft)
        /// </summary>
        public double NetVolume { get; set; }

        /// <summary>
        /// Net area (sq. ft)
        /// </summary>
        public double NetArea { get; set; }


       
    }
    

    
}
