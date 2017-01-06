using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel;
using System.Threading;
using System.Windows.Forms.DataVisualization.Charting;

namespace AnalyticalSupportData_Info
{
    public partial class Embodied_Energy_and_Carbon : Form
    {
        public Embodied_Energy_and_Carbon()
        {
            InitializeComponent();
        }


        static double M_NetVolumeValue;
        public double MaterialNetVolume_Value
        {
            set
            {
                //MessageBox.Show (value.ToString("n"));
                M_NetVolumeValue = value;
            }
        }

        static double M_NetAreaValue;
        public double MaterialNetArea_Value
        {
            set
            {
                //MessageBox.Show (value.ToString("n"));
                M_NetAreaValue = value;
            }
        }

        static string M_QuantityTypeValue;
        public string MaterialQuantityType_Value
        {
            set
            {
                //MessageBox.Show (value.ToString("n"));
                M_QuantityTypeValue = value;
            }
        }

        static string M_GifaTypeValue;
        public string MaterialGifaType_Value
        {
            set
            {
                //MessageBox.Show (value.ToString("n"));
                M_GifaTypeValue = value;
            }
        }


        static DataTable M_QuantitiesTableValue;
        public DataTable MaterialQuantitiesTable_Value
        {
            set
            {
                //MessageBox.Show (value.ToString("n"));
                M_QuantitiesTableValue = value;
            }
        }




        private void NRMRadioButton_CheckedChanged(object sender, EventArgs e)
        {

            if (sender == NRMRadioButton)
            {

                UNICLASSRadioButton.Checked = false;
                SMM7RadioButton.Checked = false;
                CESMMRadioButton.Checked = false;


                try
                {

                    //MessageBox.Show(GetElementTypeName() + " - " + M_QuantityTypeValue + ": " + M_NetVolumeValue);

                   // MessageBox.Show(M_GifaTypeValue + " - " + M_QuantityTypeValue + ": " + M_NetVolumeValue.ToString("n"));



                    GetProjectPerformanceData();

                    GifaRowNumbering();

                    //TransferQuantitiesToGifaGridView();

                    //DesignOptionDataGridView.Rows.Add(new string[] { ProjectId, DesignOptionId, TotalLCC_TextBoxValue, TotalCO2EmissionValue, TotalEcoFootprint });
                    RptGetDatasetElem();

                    TransferQuantitiesToGifaGridView();

                   


                }
                catch (Exception)
                {

                }
            }

        }



        DataTable dt = new DataTable();


        private void GetProjectPerformanceData()
        {
            
            //string ExcelInputFolder = "C:\\Users\\p0077247\\documents\\Visual Studio 2010\\Projects\\Embodied Carbon Analysis\\SMMTEMplate.xlsx";
            //string ExcelInputFolder2 = "C:\\Users\\p0077247\\documents\\Visual Studio 2010\\Projects\\Embodied Carbon Analysis\\NRMTemplate2.xlsx";
            
            string ExcelInputFolder = "";      
            string ExcelInputFolder2 = "";
            string ExcelfilePath = "";
            string ExcelfilePath2 = "";         
           
            //string dir = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "SMMTEMplate");
            string dir0 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            

            ExcelInputFolder = Path.GetFullPath(Path.Combine(dir0, @"../../")) + "SMMTEMplate.xlsx";
            ExcelInputFolder2 = Path.GetFullPath(Path.Combine(dir0, @"../../")) + "NRMTemplate2.xlsx";

            if (File.Exists(ExcelInputFolder) && File.Exists(ExcelInputFolder2))
            {
               ExcelfilePath = ExcelInputFolder;
               ExcelfilePath2 = ExcelInputFolder2;
            }
            else // file(s) not found
            {
                MessageBox.Show ("One or more resource files (SMMTEmplate, NRMTemplate2) not found in BEECE folder!");// file not found
            }

              

            //string ExfilePath = "source" ;
            if (ExcelfilePath != null)
            {
                getExcelData(ExcelfilePath);
                converToCSV(0);
                //MessageBox.Show("Oti 1");
            }

            if (ExcelfilePath2 != null)
            {

                getExcelData(ExcelfilePath2);
                converToCSV2(0);
                // MessageBox.Show("Oti 2");
            }

            else
            {
                MessageBox.Show(" File open ", "Error",
                                   MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;

            }

        }

        


        //protected static string GetSolutionFSPath()
        //{
        //    return System.IO.Directory.GetParent(System.IO.Directory.GetCurrentDirectory()).Parent.Parent.FullName;
        //}

        //protected static string GetProjectFSPath()        {
            
        //    return String.Format("{0}\\{1}", GetSolutionFSPath(), System.Reflection.Assembly.GetExecutingAssembly().GetName().Name);
        //}




        FileStream stream = null;

        DataSet result = new DataSet();
        //DataSet result2 = new DataSet();

        string[] SheetNames;

        private void getExcelData(string file)
        {

            if (file.EndsWith(".xlsx"))
            {
                // Reading from a binary Excel file (format; *.xlsx)

                //FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read);

                stream = File.Open(file, FileMode.Open, FileAccess.Read);

                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                //IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(streamReader);

                result = excelReader.AsDataSet();
                excelReader.Close();
            }

            if (file.EndsWith(".xls"))
            {

                // Reading from a binary Excel file ('97-2003 format; *.xls)
                FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                result = excelReader.AsDataSet();
                excelReader.Close();
            }

            //int[] NoOfSheet = new int [result.Tables.Count];

            SheetNames = new string[result.Tables.Count];

            List<string> items = new List<string>();

            for (int i = 0; i < result.Tables.Count; i++)
            {
                items.Add(result.Tables[i].TableName.ToString());

                SheetNames[i] = result.Tables[i].TableName.ToString();
               
            }

          



        }




        string outputCSV;
        string outputCSV2;

        DataTable[] dtM = null;
        

        private void converToCSV(int ind)
        {
            // sheets in excel file becomes tables in dataset
            //result.Tables[0].TableName.ToString(); // to get sheet name (table name)

            dtM = new DataTable[SheetNames.Length];

            //DataTable dt = new DataTable();

            for (int j = 0; j < SheetNames.Length; j++)
            {
                string a = "";
                int row_no = 0;

                while (row_no < result.Tables[j].Rows.Count)
                {
                    for (int i = 0; i < result.Tables[j].Columns.Count; i++)
                    {
                        a += result.Tables[j].Rows[row_no][i].ToString() + ",";
                    }
                    row_no++;
                    a += "\n";
                }
                
                //string MainOutputFolder = "C:\\Users\\p0077247\\documents\\Visual Studio 2010\\Projects\\Embodied Carbon Analysis\\ExcelTemplateFolder"; // define your own filepath & filename
                

                //string outputFolder = MainOutputFolder; //+ "\\" + checkedProject; ; // define your own filepath & filename




                string MainOutputFolder = "";
                string outputFolder = "";
                //string ExcelfilePath = "";
                //string ExcelfilePath2 = "";

                //string dir = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "SMMTEMplate");
                string dir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);


                MainOutputFolder = Path.GetFullPath(Path.Combine(dir1, @"../../")) + "ExcelTemplateFolder";

                if (Directory.Exists(MainOutputFolder))
                {

                    outputFolder = MainOutputFolder;
                    
                }
                else // file(s) not found
                {
                    MessageBox.Show("ExcelTemplateFolder not found in BEECE folder!");// file not found
                }






                //string output = outputFolder + "\\" + comboBox1.Text + ".csv";
                System.IO.Directory.CreateDirectory(outputFolder);
                //for (int i = 0; i < SheetNames.Length; i++)
                //{
                outputCSV = outputFolder + "\\" + SheetNames[j] + ".csv";

                //string output = outputFolder + "\\" + comboBox1.Text + ".csv";

                StreamWriter csv = new StreamWriter(@outputCSV, false);
                csv.Write(a);
                csv.Close();

                // Copy csv file to dataTable

                string[] Lines = File.ReadAllLines(outputCSV);
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

                dtM[j] = dt;


                //GetProjectPerformanceData();

            }



            ////////////////
            //////////////////

            //Put here

            GifaDataGridView.AutoGenerateColumns = false;

            GifaDataGridView.DataSource = dtM[0];

            NumberCol.DataPropertyName = "Number";
            ItemCol.DataPropertyName = "Item";
            TreeItemCol.DataPropertyName = "Tree Item";
            MaterialDescriptionCol.DataPropertyName = "Material Description";

            

            return;
        }

        DataTable[] dtM2 = null;

        private void converToCSV2(int ind)
        {
            // sheets in excel file becomes tables in dataset
            //result.Tables[0].TableName.ToString(); // to get sheet name (table name)

            dtM2 = new DataTable[SheetNames.Length];

            //DataTable dt = new DataTable();

            for (int j = 0; j < SheetNames.Length; j++)
            {
                string a = "";
                int row_no = 0;

                while (row_no < result.Tables[j].Rows.Count)
                {
                    for (int i = 0; i < result.Tables[j].Columns.Count; i++)
                    {
                        a += result.Tables[j].Rows[row_no][i].ToString() + ",";
                    }
                    row_no++;
                    a += "\n";
                }
                
                //string MainOutputFolder = "C:\\Users\\p0077247\\documents\\Visual Studio 2010\\Projects\\Embodied Carbon Analysis\\ExcelTemplateFolder2"; // define your own filepath & filename
               

                //string outputFolder2 = MainOutputFolder; //+ "\\" + checkedProject; ; // define your own filepath & filename



                string MainOutputFolder = "";
                string outputFolder2 = "";
                //string ExcelfilePath = "";
                //string ExcelfilePath2 = "";

                //string dir = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "SMMTEMplate");
                string dir2 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);


                MainOutputFolder = Path.GetFullPath(Path.Combine(dir2, @"../../")) + "ExcelTemplateFolder2";

                if (Directory.Exists(MainOutputFolder))
                {

                    outputFolder2 = MainOutputFolder;

                }
                else // file(s) not found
                {
                    MessageBox.Show("ExcelTemplateFolder2 not found in BEECE folder!");// file not found
                }





                System.IO.Directory.CreateDirectory(outputFolder2);
               
                outputCSV2 = outputFolder2 + "\\" + SheetNames[j] + ".csv";

               

                StreamWriter csv = new StreamWriter(@outputCSV2, false);
                csv.Write(a);
                csv.Close();

                string[] Lines = File.ReadAllLines(outputCSV2);
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

                dtM2[j] = dt;


                //GetProjectPerformanceData();

            }



          


            return;
        }



        private void converToCSV3(int ind)
        {
            // sheets in excel file becomes tables in dataset
            //result.Tables[0].TableName.ToString(); // to get sheet name (table name)

            dtM2 = new DataTable[SheetNames.Length];

            // dtM2 = treeViewItemsTable;

            //DataTable dt = new DataTable();

            for (int j = 0; j < SheetNames.Length; j++)
            {
                string a = "";
                int row_no = 0;

                while (row_no < treeViewItemsTable.Rows.Count)
                {
                    for (int i = 0; i < treeViewItemsTable.Columns.Count; i++)
                    {
                        a += treeViewItemsTable.Rows[row_no][i].ToString() + ",";
                    }
                    row_no++;
                    a += "\n";
                }


               
                //string MainOutputFolder = "C:\\Users\\p0077247\\documents\\Visual Studio 2010\\Projects\\Embodied Carbon Analysis\\ExcelTemplateFolder3"; // define your own filepath & filename
              

               // string outputFolder3 = MainOutputFolder; //+ "\\" + checkedProject; ; // define your own filepath & filename



                string MainOutputFolder = "";
                string outputFolder3 = "";
                //string ExcelfilePath = "";
                //string ExcelfilePath2 = "";

                //string dir = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "SMMTEMplate");
                string dir3 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);


                MainOutputFolder = Path.GetFullPath(Path.Combine(dir3, @"../../")) + "ExcelTemplateFolder";

                if (Directory.Exists(MainOutputFolder))
                {

                    outputFolder3 = MainOutputFolder;

                }
                else // file(s) not found
                {
                    MessageBox.Show("ExcelTemplateFolder3 not found in BEECE folder!");// file not found
                }




                System.IO.Directory.CreateDirectory(outputFolder3);
                //for (int i = 0; i < SheetNames.Length; i++)
                //{
                outputCSV2 = outputFolder3 + "\\" + SheetNames[j] + ".csv";

                //string output = outputFolder + "\\" + comboBox1.Text + ".csv";

                StreamWriter csv = new StreamWriter(@outputCSV2, false);
                csv.Write(a);
                csv.Close();

                string[] Lines = File.ReadAllLines(outputCSV2);
                string[] Fields;
                Fields = Lines[0].Split(new char[] { ',' });
                int Cols = Fields.GetLength(0);

                DataTable dt = new DataTable();
                dt = treeViewItemsTable;
               

                dtM2[j] = dt;


                //GetProjectPerformanceData();

            }



         


            return;
        }


       



        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'iCEDatabaseDataSet.ICEItemTable' table. You can move, or remove it, as needed.
            this.iCEItemTableTableAdapter.Fill(this.iCEDatabaseDataSet.ICEItemTable);



        }




        ICEDatabaseDataSet.ICEItemTableRow ICEInfoRow;
        //int ICEInfoRowNumber = 0;
        int NumberID = 0;


       

        private void iCEItemTableBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.iCEItemTableBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.iCEDatabaseDataSet);

        }


        int ICEInfoRowIndex = 0;
        int GifaDataGridViewRowNo = 0;
        int GifaDataGridViewRowNo2 = 0;
        double VolumeValue = 00.0;
       // double DensityValue = 0.0;
        //double EEI_Value = 0.0; // Convert.ToDouble(ICEInfoRow.EE);
        //double ECI_Value = 0.0; //Convert.ToDouble(ICEInfoRow.EC)

        private void GifaDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            double EEI_Value = 0.0; // Convert.ToDouble(ICEInfoRow.EE);
            double ECI_Value = 0.0; //Convert.ToDouble(ICEInfoRow.EC)
            double DensityValue = 0.0;
           




            if (e.ColumnIndex == MaterialType.Index && e.RowIndex >= 0) //check if combobox column
            {
                object selectedValue = GifaDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;

                string SelItem = GifaDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].RowIndex.ToString();


                GifaDataGridViewRowNo = GifaDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].RowIndex;



                NumberID = ICEInfoRowIndex + 2;
                //NumberID = ICEInfoRowIndex;

                //int GifaDataGridViewRowNo = GifaDataGridViewRow + 1;

                //ICEInfoRow = (ICEDatabaseDataSet.ICEItemTableRow)iCEDatabaseDataSet.ICEItemTable
                ICEInfoRow = (ICEDatabaseDataSet.ICEItemTableRow)iCEDatabaseDataSet.ICEItemTable.Rows.Find(NumberID);

               

                GifaDataGridView[4, GifaDataGridViewRowNo].Value = ICEInfoRow.Density;
                GifaDataGridView[6, GifaDataGridViewRowNo].Value = ICEInfoRow.EE;
                GifaDataGridView[7, GifaDataGridViewRowNo].Value = ICEInfoRow.EC;

                //DensityValue = Convert.ToDouble(ICEInfoRow.Density);
                //EEIValue = Convert.ToDouble(ICEInfoRow.EE);
                //ECIValue = Convert.ToDouble(ICEInfoRow.EC) / 1000; // change to tonne

                //GifaDataGridView[5, ICEInfoRowIndex].Value = ICEInfoRow.EC;

                DensityValue = Convert.ToDouble(ICEInfoRow.Density);
                EEI_Value = Convert.ToDouble(ICEInfoRow.EE);
                ECI_Value = Convert.ToDouble(ICEInfoRow.EC);

                //MessageBox.Show(ECIValue.ToString());
                //MessageBox.Show("Density:  " + DensityValue.ToString());

                double EEValue = 0.0;
                double ECValue = 0.0;


               // GifaDataGridView[2, GifaDataGridViewRowNo].Value = VolumeTrialValue;


                VolumeValue = Convert.ToDouble(GifaDataGridView[2, GifaDataGridViewRowNo].Value);


                if (GifaDataGridView[2, GifaDataGridViewRowNo].Value != null)
                {
                    GifaDataGridView[5, GifaDataGridViewRowNo].Value = VolumeValue * DensityValue; // mass

                  

                    EEValue = Convert.ToDouble(GifaDataGridView[4, GifaDataGridViewRowNo].Value) *
                        Convert.ToDouble(GifaDataGridView[2, GifaDataGridViewRowNo].Value) *
                        Convert.ToDouble(GifaDataGridView[6, GifaDataGridViewRowNo].Value) / 1000;
                    //DensityValue * VolumeTrialValue * EEIValue;


                    ECValue = Convert.ToDouble(GifaDataGridView[4, GifaDataGridViewRowNo].Value) *
                        Convert.ToDouble(GifaDataGridView[2, GifaDataGridViewRowNo].Value) *
                        Convert.ToDouble(GifaDataGridView[7, GifaDataGridViewRowNo].Value) / 1000 ;



                    GifaDataGridView[8, GifaDataGridViewRowNo].Value = EEValue.ToString();
                    GifaDataGridView[9, GifaDataGridViewRowNo].Value = ECValue.ToString();
                }

                //MessageBox.Show(VolumeTrialValue.ToString());

            }


            // if (e.ColumnIndex == 2)

            if (e.ColumnIndex == 2 && e.RowIndex >= 0)
            {
               
                

                GifaDataGridViewRowNo2 = GifaDataGridView.Rows[e.RowIndex].Cells[2].RowIndex;

                //GifaDataGridViewRowNo = GifaDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].RowIndex;



               

                double EEValue = 0.0;
                double ECValue = 0.0;
                //GifaDataGridView[2, GifaDataGridViewRowNo].Value = VolumeTrialValue;


                VolumeValue = Convert.ToDouble(GifaDataGridView[2, GifaDataGridViewRowNo2].Value);

                
               // (gridRow.Cells[i1].Value == null ? DBNull.Value : gridRow.Cells[i1].Value)

                if (GifaDataGridView[5, GifaDataGridViewRowNo2].Value != null)
                {

                GifaDataGridView[5, GifaDataGridViewRowNo2].Value = VolumeValue * DensityValue;


                EEValue = Convert.ToDouble(GifaDataGridView[4, GifaDataGridViewRowNo2].Value) *
                    Convert.ToDouble(GifaDataGridView[2, GifaDataGridViewRowNo2].Value) *
                    Convert.ToDouble(GifaDataGridView[6, GifaDataGridViewRowNo2].Value) /1000;
                //DensityValue * VolumeTrialValue * EEIValue;


                ECValue = Convert.ToDouble(GifaDataGridView[4, GifaDataGridViewRowNo2].Value) *
                    Convert.ToDouble(GifaDataGridView[2, GifaDataGridViewRowNo2].Value) *
                    Convert.ToDouble(GifaDataGridView[7, GifaDataGridViewRowNo2].Value) / 1000;
                //ECValue = DensityValue * VolumeTrialValue * ECIValue;


                GifaDataGridView[8, GifaDataGridViewRowNo2].Value = EEValue.ToString();
                GifaDataGridView[9, GifaDataGridViewRowNo2].Value = ECValue.ToString();


                }

            }








        }





        private void GifaDataGridView_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (GifaDataGridView.IsCurrentCellDirty)
            {
                GifaDataGridView.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }



        private void GifaDataGridView_EditingControlShowing(object sender,
        DataGridViewEditingControlShowingEventArgs e)
        {

            if (GifaDataGridView.CurrentCell.ColumnIndex == 3)
            {
                ComboBox combo = e.Control as ComboBox;
                if (combo != null)
                {
                    // Remove an existing event-handler, if present, to avoid 
                    // adding multiple handlers when the editing control is reused.
                    combo.SelectedIndexChanged -=
                        new EventHandler(ComboB_SelectedIndexChanged);

                    // Add the event handler. 
                    combo.SelectedIndexChanged +=
                        new EventHandler(ComboB_SelectedIndexChanged);

                }

                //MessageBox.Show(ComboBoxselectedIndex.ToString());
            }

            //GifaDataGridView.Invalidate();
        }


        public void GifaRowNumbering()
        {

            //Number rows on rowHeader
            int rowNumber = 1;
            foreach (DataGridViewRow row in GifaDataGridView.Rows)
            {
                if (row.IsNewRow) continue;
                //row.HeaderCell.Value = rowNumber;
                row.HeaderCell.Value = rowNumber.ToString();

                rowNumber = rowNumber + 1;
            }

            //Resize rowHeader column
            GifaDataGridView.AutoResizeRowHeadersWidth(
                DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);


        }


        public void GifaSummRowNumbering2()
        {

            //Number rows on rowHeader
            int rowNumber = 1;
            foreach (DataGridViewRow row in GifaSummaryDataGridView.Rows)
            {
                if (row.IsNewRow) continue;
                //row.HeaderCell.Value = rowNumber;
                row.HeaderCell.Value = rowNumber.ToString();

                rowNumber = rowNumber + 1;
            }

            //Resize rowHeader column
            GifaSummaryDataGridView.AutoResizeRowHeadersWidth(
                DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);


        }

       



        int ComboBoxselectedIndex = 0;

        private void ComboB_SelectedIndexChanged(object sender, EventArgs e)
        {
            //((ComboBox)sender).BackColor = (Color)((ComboBox)sender).SelectedItem;
            // MessageBox.Show("oti");
            ComboBoxselectedIndex = ((ComboBox)sender).SelectedIndex;
            //MessageBox.Show("Selected Index = " + ComboBoxselectedIndex);

            //GifaDataGridView.Invalidate();
            //ICEInfoRowIndex = GifaDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].RowIndex;
            if (ComboBoxselectedIndex >= 0)
            {

                ICEInfoRowIndex = ComboBoxselectedIndex;


            }


            //GifaDataGridView.Invalidate();


        }





        /// <summary>
        /// 
        ///
        /// </summary>

        //DataTable dt3 = null;
        DataTable dt2 = null;





        int NumOfElements = 0;
        int NumOfGifas = 0;

        private void button4_Click(object sender, EventArgs e)
        {

            //RptGetDatasetElem();

            CreateMatirixByTranspose();

            FillSummaryTable();

            PlotGenericCharts();

        }

        int SummaryRun = 1;

        public void FillSummaryTable()
        {
            //SummaryRun = 1;

           // dt2 = null; 

            // RptGetDatasetElem();

            dt2 = Matrix_mTable;

            NumOfElements = dt2.Columns.Count;

            NumOfGifas = dt2.Rows.Count + 1;

            //MessageBox.Show(NumOfGifas.ToString());
            //MessageBox.Show(NumOfElements.ToString());



            double[][] m = MatrixCreate(NumOfGifas, NumOfElements);


            double[][] arr = new double[NumOfGifas][];
            double[] myArrayList0 = new double[NumOfElements];
            double[] myArrayList1 = new double[NumOfElements];
            double[] myArrayList2 = new double[NumOfElements];
            double[] myArrayList3 = new double[NumOfElements];
            double[] myArrayList4 = new double[NumOfElements];
            double[] myArrayList5 = new double[NumOfElements];
            double[] myArrayList6 = new double[NumOfElements];
            double[] myArrayList7 = new double[NumOfElements];


            double[] myArrayList8 = new double[NumOfElements];
            double[] myArrayList9 = new double[NumOfElements];
           double[] myArrayList10 = new double[NumOfElements];
            //double[] myArrayList10= new double[NumOfElements];


            //ArrayList myArrayList = new ArrayList(2000);

            //for (int i = 1; i < dt2.Rows.Count - 1; i++)
            //{
            for (int jj = 1; jj < dt2.Columns.Count; jj++)
            {

                //if (dt2.Rows[0][j].ToString() != "SubstructureTabPage")
                //{
                myArrayList0[jj] = Convert.ToDouble(dt2.Rows[0][jj]);
                myArrayList1[jj] = Convert.ToDouble(dt2.Rows[1][jj]);
                myArrayList2[jj] = Convert.ToDouble(dt2.Rows[2][jj]);
                myArrayList3[jj] = Convert.ToDouble(dt2.Rows[3][jj]);
                myArrayList4[jj] = Convert.ToDouble(dt2.Rows[4][jj]);
                myArrayList5[jj] = Convert.ToDouble(dt2.Rows[5][jj]);
                myArrayList6[jj] = Convert.ToDouble(dt2.Rows[6][jj]);
                myArrayList7[jj] = Convert.ToDouble(dt2.Rows[7][jj]);
                myArrayList8[jj] = Convert.ToDouble(dt2.Rows[8][jj]);
                //myArrayList9[j] = Convert.ToDouble(dt2.Rows[9][j]);
                //myArrayList10[j] = Convert.ToDouble(dt2.Rows[10][j]);
                
                //richTextBox1.Text = myArrayList5[jj].ToString();
                //richTextBox2.Text = myArrayList6[jj].ToString();
                //richTextBox3.Text = myArrayList7[jj].ToString();
                //richTextBox4.Text = myArrayList8[jj].ToString();

               // MessageBox.Show(myArrayList0[jj].ToString());
                //}

                // richTextBox4.Text = myArrayList[j].ToString();
            }

            for (int jj = 1; jj < ModifiedTransposedTable.Columns.Count; jj++)
            {

                //if (dt2.Rows[0][j].ToString() != "SubstructureTabPage")
                //{

                //myArrayList8[jj] = Convert.ToDouble(ModifiedTransposedTable.Rows[6][jj]);
                //myArrayList9[jj] = Convert.ToDouble(ModifiedTransposedTable.Rows[7][jj]);

                myArrayList9[jj] = Convert.ToDouble(ModifiedTransposedTable.Rows[5][jj]);
                myArrayList10[jj] = Convert.ToDouble(ModifiedTransposedTable.Rows[6][jj]);

              
            }

            

            m[0] = myArrayList0;
            m[1] = myArrayList1;
            //arr[1] = myArrayList;
            m[2] = myArrayList2;
            m[3] = myArrayList3;
            m[4] = myArrayList4;
            m[5] = myArrayList5;
            m[6] = myArrayList6;
            m[7] = myArrayList7;
            m[8] = myArrayList8;



            //Console.WriteLine("Matrix m = \n" + MatrixAsString(m));
            //richTextBox1.Text = MatrixAsString(m);

            MatrixAsString(m);


            //Console.WriteLine("Matrix n1 = \n" + MatrixAsString(n1));

            //double[] n2 = new double[] { 3.0, 1.0, 2.0, 5.0, 5.0 };
            double[] n2 = new double[NumOfElements];
            double[] n3 = new double[NumOfElements];

            n2 = myArrayList9;
            n3 = myArrayList10;

            //Console.WriteLine("MatrixVectorProduct = \n" + VectorAsString(MatrixVectorProduct(m, n2)));

            //richTextBox2.Text = MatrixAsString(n1);

            ///

            VectorAsString(n2);
            VectorAsStringGIFA(MatrixVectorProduct(m, n2));

           


            GifaDataGridView[10, Gifa0_RowToColIndex].Value = Gifa0_Values;
            GifaDataGridView[10, Gifa1_RowToColIndex].Value = Gifa1_Values;
            GifaDataGridView[10, Gifa2_RowToColIndex].Value = Gifa2_Values;
            GifaDataGridView[10, Gifa3_RowToColIndex].Value = Gifa3_Values;
            GifaDataGridView[10, Gifa4_RowToColIndex].Value = Gifa4_Values;
            GifaDataGridView[10, Gifa5_RowToColIndex].Value = Gifa5_Values;
            GifaDataGridView[10, Gifa6_RowToColIndex].Value = Gifa6_Values;
            GifaDataGridView[10, Gifa7_RowToColIndex].Value = Gifa7_Values;
            GifaDataGridView[10, Gifa8_RowToColIndex].Value = Gifa8_Values;

            GifaDataGridView[10, GifaSum_RowToColIndex + 1].Value = GifaSum_Values;
            //GifaDataGridView[10, GifaSum_RowToColIndex + 1].Value = GifaSum_Values;

            ///
           // MessageBox.Show("G1: -  " + Gifa2_Values);

            ///

            VectorAsStringGIFA(MatrixVectorProduct(m, n3));

            // richTextBox7.Text = VectorAsStringGIFA(MatrixVectorProduct(m, n3));

            //richTextBox8.Text = TotalOfGifaValues.ToString();

            GifaDataGridView[11, Gifa0_RowToColIndex].Value = Gifa0_Values;
            GifaDataGridView[11, Gifa1_RowToColIndex].Value = Gifa1_Values;
            GifaDataGridView[11, Gifa2_RowToColIndex].Value = Gifa2_Values;
            GifaDataGridView[11, Gifa3_RowToColIndex].Value = Gifa3_Values;
            GifaDataGridView[11, Gifa4_RowToColIndex].Value = Gifa4_Values;
            GifaDataGridView[11, Gifa5_RowToColIndex].Value = Gifa5_Values;
            GifaDataGridView[11, Gifa6_RowToColIndex].Value = Gifa6_Values;
            GifaDataGridView[11, Gifa7_RowToColIndex].Value = Gifa7_Values;
            GifaDataGridView[11, Gifa8_RowToColIndex].Value = Gifa8_Values;

            //MessageBox.Show(Gifa7_Values.ToString());
            GifaDataGridView[11, GifaSum_RowToColIndex + 1].Value = GifaSum_Values;
            GifaDataGridView[1, GifaSum_RowToColIndex + 1].Value = "TOTAL";
            ///
            //MessageBox.Show("G2: -  " + Gifa2_Values);

           


            GifaSummaryDataGridView.Rows.Add(new string[] { "", "SUMMARY RUN  " + SummaryRun, "", "" });

            GifaSummaryDataGridView.Rows.Add(new string[] { GifaDataGridView[0, Gifa0_RowToColIndex].Value.ToString(), 
                GifaDataGridView[1, Gifa0_RowToColIndex].Value.ToString(), GifaDataGridView[10, Gifa0_RowToColIndex].Value.ToString(), GifaDataGridView[11, Gifa0_RowToColIndex].Value.ToString() });

            GifaSummaryDataGridView.Rows.Add(new string[] { GifaDataGridView[0, Gifa1_RowToColIndex].Value.ToString(), 
                GifaDataGridView[1, Gifa1_RowToColIndex].Value.ToString(), GifaDataGridView[10, Gifa1_RowToColIndex].Value.ToString(), GifaDataGridView[11, Gifa1_RowToColIndex].Value.ToString() });

            GifaSummaryDataGridView.Rows.Add(new string[] { GifaDataGridView[0, Gifa2_RowToColIndex].Value.ToString(), 
                GifaDataGridView[1, Gifa2_RowToColIndex].Value.ToString(), GifaDataGridView[10, Gifa2_RowToColIndex].Value.ToString(), GifaDataGridView[11, Gifa2_RowToColIndex].Value.ToString() });

            GifaSummaryDataGridView.Rows.Add(new string[] { GifaDataGridView[0, Gifa3_RowToColIndex].Value.ToString(), 
                GifaDataGridView[1, Gifa3_RowToColIndex].Value.ToString(), GifaDataGridView[10, Gifa3_RowToColIndex].Value.ToString(), GifaDataGridView[11, Gifa3_RowToColIndex].Value.ToString() });

            GifaSummaryDataGridView.Rows.Add(new string[] { GifaDataGridView[0, Gifa4_RowToColIndex].Value.ToString(), 
                GifaDataGridView[1, Gifa4_RowToColIndex].Value.ToString(), GifaDataGridView[10, Gifa4_RowToColIndex].Value.ToString(), GifaDataGridView[11, Gifa4_RowToColIndex].Value.ToString() });

            GifaSummaryDataGridView.Rows.Add(new string[] { GifaDataGridView[0, Gifa5_RowToColIndex].Value.ToString(), 
                GifaDataGridView[1, Gifa5_RowToColIndex].Value.ToString(), GifaDataGridView[10, Gifa5_RowToColIndex].Value.ToString(), GifaDataGridView[11, Gifa5_RowToColIndex].Value.ToString() });

            GifaSummaryDataGridView.Rows.Add(new string[] { GifaDataGridView[0, Gifa6_RowToColIndex].Value.ToString(), 
                GifaDataGridView[1, Gifa6_RowToColIndex].Value.ToString(), GifaDataGridView[10, Gifa6_RowToColIndex].Value.ToString(), GifaDataGridView[11, Gifa6_RowToColIndex].Value.ToString() });

            GifaSummaryDataGridView.Rows.Add(new string[] { GifaDataGridView[0, Gifa7_RowToColIndex].Value.ToString(), 
                GifaDataGridView[1, Gifa7_RowToColIndex].Value.ToString(), GifaDataGridView[10, Gifa7_RowToColIndex].Value.ToString(), GifaDataGridView[11, Gifa7_RowToColIndex].Value.ToString() });
            
            GifaSummaryDataGridView.Rows.Add(new string[] { GifaDataGridView[0, Gifa8_RowToColIndex].Value.ToString(), 
                GifaDataGridView[1, Gifa8_RowToColIndex].Value.ToString(), GifaDataGridView[10, Gifa8_RowToColIndex].Value.ToString(), GifaDataGridView[11, Gifa8_RowToColIndex].Value.ToString() });

            GifaSummaryDataGridView.Rows.Add(new string[] { "", 
                GifaDataGridView[1, GifaSum_RowToColIndex + 1].Value.ToString(), GifaDataGridView[10, GifaSum_RowToColIndex + 1].Value.ToString(), GifaDataGridView[11, GifaSum_RowToColIndex + 1].Value.ToString() });

            GifaSummaryDataGridView.Rows.Add(new string[] { "", "", "", "" });

            GifaSummRowNumbering2();

            SummaryRun = SummaryRun + 1;

        }


        //////////////


      

        DataTable Matrix_mTable = null;
        DataTable ModifiedTransposedTable = null;

        int Gifa0_RowToColIndex = 0;
        int Gifa1_RowToColIndex = 0;
        int Gifa2_RowToColIndex = 0;
        int Gifa3_RowToColIndex = 0;
        int Gifa4_RowToColIndex = 0;
        int Gifa5_RowToColIndex = 0;
        int Gifa6_RowToColIndex = 0;
        int Gifa7_RowToColIndex = 0;
        int Gifa8_RowToColIndex = 0;
        int GifaSum_RowToColIndex = 0;


        // Copy manual input data in DataGridView to DataTable and create an Xml file 
        private void RptGetDatasetElem()
        {

            //DataSet ds = new DataSet();
            ////ds.DataSetName = "MyData2";

            //ds.Tables.Add(ReportElementsInfo());

            ////dataGridView3.DataSource = ds.Tables[0];


            //DataTable inputTable = ds.Tables[0];
            //    // Table shown in Figure 1.1

            //DataTable transposedTable = GenerateTransposedTable(inputTable);



            if (dtM2 != null)
            {
                //         private void PopulateTreeView_Load(object sender, EventArgs e)
                //{
                DataTable table = new DataTable();
                //table.Columns.Add("Level");
                //table.Columns.Add("Data");
                //if (dtM2[0] != null)
                table = dtM2[0];


                //////////////////////////////////////////////////////////////


                //TreeNode lastNode = null;
                TreeNode lastNode = new TreeNode();
                //TreeNode lastNode2 = new TreeNode();  

                for (int i = 0; i < table.Rows.Count; i++)
                {

                    string rowContent = table.Rows[i]["Group1"].ToString();






                    if (table.Rows[i]["Group1"] != DBNull.Value)
                    {
                        TreeNode newNode = new TreeNode((string)table.Rows[i]["Number"] + " - " + (string)table.Rows[i]["Group1"]);

                        //TreeNode newNode = new TreeNode((string)table.Rows[i]["Group1"]);

                        if (i == 0)

                            treeView2.Nodes.Add(newNode);

                        else
                        {


                            DataRow tableRow = table.NewRow();

                            //int currentLevel1 = tableRow.Table.Columns["Group1"].Ordinal;


                            string currentRowLevel1 = table.Rows[i]["Group1"].ToString();
                            string lastRowLevel1 = table.Rows[i - 1]["Group1"].ToString();

                            string currentRowLevel2 = table.Rows[i]["Group2"].ToString();
                            string lastRowLevel2 = table.Rows[i - 1]["Group2"].ToString();


                            //TreeNode lastNode2 = new TreeNode();
                            //TreeNode currentNode; // = new TreeNode();

                            if (currentRowLevel1 == lastRowLevel1)
                            {
                                if (lastRowLevel2 == currentRowLevel2)
                                {
                                    // MessageBox.Show(newNode2.ToString());
                                }
                                else
                                {

                                    int table_RowIndex0 = 0;
                                    int table_RowIndex1 = 0;
                                    int table_RowIndex2 = 0;
                                    int table_RowIndex3 = 0;
                                    int table_RowIndex4 = 0;
                                    int table_RowIndex5 = 0;
                                    int table_RowIndex6 = 0;
                                    int table_RowIndex7 = 0;
                                    int table_RowIndex8 = 0;
                                    int table_RowIndexSum = 0;

                                    // string expression = " Item = 'Substructure'";
                                    string Number0_Expression = " Number = '0'";
                                    DataRow[] table0_foundRows = table.Select(Number0_Expression); ;
                                    table_RowIndex0 = table.Rows.IndexOf(table0_foundRows[0]);

                                    string Number1_Expression = " Number = '1'";
                                    DataRow[] table1_foundRows = table.Select(Number1_Expression); ;
                                    table_RowIndex1 = table.Rows.IndexOf(table1_foundRows[0]);


                                    string Number2_Expression = " Number = '2'";
                                    DataRow[] table2_foundRows = table.Select(Number2_Expression); ;
                                    table_RowIndex2 = table.Rows.IndexOf(table2_foundRows[0]);


                                    string Number3_Expression = " Number = '3'";
                                    DataRow[] table3_foundRows = table.Select(Number3_Expression); ;
                                    table_RowIndex3 = table.Rows.IndexOf(table3_foundRows[0]);


                                    string Number4_Expression = " Number = '4'";
                                    DataRow[] table4_foundRows = table.Select(Number4_Expression); ;
                                    table_RowIndex4 = table.Rows.IndexOf(table4_foundRows[0]);

                                    string Number5_Expression = " Number = '5'";
                                    DataRow[] table5_foundRows = table.Select(Number5_Expression); ;
                                    table_RowIndex5 = table.Rows.IndexOf(table5_foundRows[0]);

                                    string Number6_Expression = " Number = '6'";
                                    DataRow[] table6_foundRows = table.Select(Number6_Expression); ;
                                    table_RowIndex6 = table.Rows.IndexOf(table6_foundRows[0]);

                                    string Number7_Expression = " Number = '7'";
                                    DataRow[] table7_foundRows = table.Select(Number7_Expression); ;
                                    table_RowIndex7 = table.Rows.IndexOf(table7_foundRows[0]);

                                    string Number8_Expression = " Number = '8'";
                                    DataRow[] table8_foundRows = table.Select(Number8_Expression); ;
                                    table_RowIndex8 = table.Rows.IndexOf(table8_foundRows[0]);

                                    string NumberSum_Expression = " Number = '8.8.3.1'";
                                    DataRow[] tableSum_foundRows = table.Select(NumberSum_Expression); ;
                                    table_RowIndexSum = table.Rows.IndexOf(tableSum_foundRows[0]);





                                    ///////////////////////////////////////////////////////////////////////////////////



                                    if (i > 0)
                                    {

                                        if (i < table_RowIndex1)
                                        // if (j < table_RowIndex1)
                                        {


                                            for (int j = 0; j < table_RowIndex1; j++)
                                            {
                                                //string currentRowLevel10 = table.Rows[j]["Group1"].ToString();
                                                //string lastRowLevel10 = table.Rows[table_RowIndex0]["Group1"].ToString();
                                                if (j > 0)
                                                {
                                                    //string currentRowLevel20 = table.Rows[j]["Group2"].ToString();
                                                    //string lastRowLevel20 = table.Rows[j - 2]["Group2"].ToString();

                                                    string currentRowLevel10 = table.Rows[j]["Group1"].ToString();
                                                    string lastRowLevel10 = table.Rows[table_RowIndex0]["Group1"].ToString();

                                                    string currentRowLevel20 = table.Rows[j]["Group2"].ToString();
                                                    string lastRowLevel20 = table.Rows[j - 1]["Group2"].ToString();

                                                    string currentRowLevel20_Num = table.Rows[j]["Number"].ToString();


                                                    if (lastRowLevel20 != currentRowLevel20)
                                                    {

                                                        TreeNode childNode = new TreeNode(currentRowLevel20_Num + " - " + currentRowLevel20);
                                                        //TreeNode childNode = new TreeNode(currentRowLevel20);
                                                        //TreeNode childNode = new TreeNode(currentRowLevel20_Num);

                                                        if (lastRowLevel10 == currentRowLevel10)
                                                        {
                                                            lastNode.Nodes.Add(childNode);


                                                            for (int jj = 0; jj < table_RowIndex1; jj++)
                                                            {
                                                                if (jj > 0) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                {
                                                                    string currentRowLevel20_2 = table.Rows[jj]["Group2"].ToString();
                                                                    string lastRowLevel20_2 = table.Rows[table_RowIndex0]["Group2"].ToString();

                                                                    string currentRowLevel30 = table.Rows[jj]["Group3"].ToString();
                                                                    string lastRowLevel30 = table.Rows[jj - 1]["Group3"].ToString();

                                                                    string currentRowLevel30_Num = table.Rows[jj]["Number"].ToString();

                                                                    if (currentRowLevel30 != "")
                                                                    {
                                                                        TreeNode childNode3 = new TreeNode(currentRowLevel30_Num + " - " + currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30_Num);

                                                                        if (lastRowLevel30 != currentRowLevel30)
                                                                        {
                                                                            if (currentRowLevel20_2 == currentRowLevel20)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                            {
                                                                                childNode.Nodes.Add(childNode3);


                                                                                for (int jjj = 0; jjj < table_RowIndex1; jjj++)
                                                                                {
                                                                                    if (jjj > 0) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                    {
                                                                                        string currentRowLevel30_2 = table.Rows[jjj]["Group3"].ToString();
                                                                                        string lastRowLevel30_2 = table.Rows[table_RowIndex0]["Group3"].ToString();

                                                                                        string currentRowLevel40 = table.Rows[jjj]["Group4"].ToString();
                                                                                        string lastRowLevel40 = table.Rows[jjj - 1]["Group4"].ToString();

                                                                                        string currentRowLevel40_Num = table.Rows[jjj]["Number"].ToString();

                                                                                        if (currentRowLevel40 != "")
                                                                                        {
                                                                                            TreeNode childNode4 = new TreeNode(currentRowLevel40_Num + " - " + currentRowLevel40);
                                                                                            // TreeNode childNode4 = new TreeNode( currentRowLevel40);
                                                                                            //TreeNode childNode4 = new TreeNode(currentRowLevel40_Num);

                                                                                            if (lastRowLevel40 != currentRowLevel40)
                                                                                            {
                                                                                                if (currentRowLevel30_2 == currentRowLevel30)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                {
                                                                                                    childNode3.Nodes.Add(childNode4);


                                                                                                    for (int jjjj = 0; jjjj < table_RowIndex1; jjjj++)
                                                                                                    {
                                                                                                        if (jjjj > 0) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                        {
                                                                                                            string currentRowLevel40_2 = table.Rows[jjjj]["Group4"].ToString();
                                                                                                            string lastRowLevel40_2 = table.Rows[table_RowIndex0]["Group4"].ToString();

                                                                                                            string currentRowLevel50 = table.Rows[jjjj]["Group5"].ToString();
                                                                                                            string lastRowLevel50 = table.Rows[jjjj - 1]["Group5"].ToString();

                                                                                                            string currentRowLevel50_Num = table.Rows[jjjj]["Number"].ToString();

                                                                                                            if (currentRowLevel50 != "")
                                                                                                            {
                                                                                                                TreeNode childNode5 = new TreeNode(currentRowLevel50_Num + " - " + currentRowLevel50);
                                                                                                                //TreeNode childNode5 = new TreeNode(currentRowLevel50);
                                                                                                                //TreeNode childNode5 = new TreeNode(currentRowLevel50_Num);

                                                                                                                if (lastRowLevel50 != currentRowLevel50)
                                                                                                                {
                                                                                                                    if (currentRowLevel40_2 == currentRowLevel40)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                                    {
                                                                                                                        childNode4.Nodes.Add(childNode5);



                                                                                                                    }
                                                                                                                    //

                                                                                                                }
                                                                                                            }
                                                                                                        }
                                                                                                    }



                                                                                                }
                                                                                                //

                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                            //

                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            // }
                                                        }

                                                    }
                                                }
                                            }
                                        }


                                        if (i >= table_RowIndex1 && i < table_RowIndex2)
                                        {

                                            for (int j = table_RowIndex1; j < table_RowIndex2; j++)
                                            {

                                                if (j > table_RowIndex1)
                                                {
                                                    //string currentRowLevel20 = table.Rows[j]["Group2"].ToString();
                                                    //string lastRowLevel20 = table.Rows[j - 2]["Group2"].ToString();

                                                    string currentRowLevel10 = table.Rows[j]["Group1"].ToString();
                                                    string lastRowLevel10 = table.Rows[table_RowIndex1]["Group1"].ToString();

                                                    string currentRowLevel20 = table.Rows[j]["Group2"].ToString();
                                                    string lastRowLevel20 = table.Rows[j - 1]["Group2"].ToString();

                                                    string currentRowLevel20_Num = table.Rows[j]["Number"].ToString();

                                                    if (lastRowLevel20 != currentRowLevel20)
                                                    {

                                                        TreeNode childNode = new TreeNode(currentRowLevel20_Num + " - " + currentRowLevel20);
                                                        //TreeNode childNode = new TreeNode(currentRowLevel20);
                                                        //TreeNode childNode = new TreeNode(currentRowLevel20_Num);

                                                        if (lastRowLevel10 == currentRowLevel10)
                                                        {
                                                            lastNode.Nodes.Add(childNode);


                                                            for (int jj = table_RowIndex1; jj < table_RowIndex2; jj++)
                                                            //for (int jj = 0; jj < table_RowIndex1; jj += 2)
                                                            {
                                                                if (jj > table_RowIndex1) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                {
                                                                    string currentRowLevel20_2 = table.Rows[jj]["Group2"].ToString();
                                                                    string lastRowLevel20_2 = table.Rows[table_RowIndex0]["Group2"].ToString();

                                                                    string currentRowLevel30 = table.Rows[jj]["Group3"].ToString();
                                                                    string lastRowLevel30 = table.Rows[jj - 1]["Group3"].ToString();

                                                                    string currentRowLevel30_Num = table.Rows[jj]["Number"].ToString();

                                                                    if (currentRowLevel30 != "")
                                                                    {
                                                                        TreeNode childNode3 = new TreeNode(currentRowLevel30_Num + " - " + currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30_Num);

                                                                        if (lastRowLevel30 != currentRowLevel30)
                                                                        {
                                                                            if (currentRowLevel20_2 == currentRowLevel20)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                            {

                                                                                childNode.Nodes.Add(childNode3);


                                                                                for (int jjj = table_RowIndex1; jjj < table_RowIndex2; jjj++)
                                                                                //for (int jjj = 0; jjj < table_RowIndex1; jjj += 2)
                                                                                {
                                                                                    if (jjj > table_RowIndex1) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                    {
                                                                                        string currentRowLevel30_2 = table.Rows[jjj]["Group3"].ToString();
                                                                                        string lastRowLevel30_2 = table.Rows[table_RowIndex0]["Group3"].ToString();

                                                                                        string currentRowLevel40 = table.Rows[jjj]["Group4"].ToString();
                                                                                        string lastRowLevel40 = table.Rows[jjj - 1]["Group4"].ToString();

                                                                                        string currentRowLevel40_Num = table.Rows[jjj]["Number"].ToString();

                                                                                        if (currentRowLevel40 != "")
                                                                                        {
                                                                                            TreeNode childNode4 = new TreeNode(currentRowLevel40_Num + " - " + currentRowLevel40);
                                                                                            //TreeNode childNode4 = new TreeNode( currentRowLevel40);
                                                                                            //TreeNode childNode4 = new TreeNode(currentRowLevel40_Num);

                                                                                            if (lastRowLevel40 != currentRowLevel40)
                                                                                            {
                                                                                                if (currentRowLevel30_2 == currentRowLevel30)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                {
                                                                                                    childNode3.Nodes.Add(childNode4);



                                                                                                    for (int jjjj = table_RowIndex1; jjjj < table_RowIndex2; jjjj++)
                                                                                                    {
                                                                                                        if (jjjj > table_RowIndex1)  //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                        {
                                                                                                            string currentRowLevel40_2 = table.Rows[jjjj]["Group4"].ToString();
                                                                                                            string lastRowLevel40_2 = table.Rows[table_RowIndex0]["Group4"].ToString();

                                                                                                            string currentRowLevel50 = table.Rows[jjjj]["Group5"].ToString();
                                                                                                            string lastRowLevel50 = table.Rows[jjjj - 1]["Group5"].ToString();

                                                                                                            string currentRowLevel50_Num = table.Rows[jjjj]["Number"].ToString();

                                                                                                            if (currentRowLevel50 != "")
                                                                                                            {
                                                                                                                TreeNode childNode5 = new TreeNode(currentRowLevel50_Num + " - " + currentRowLevel50);
                                                                                                                //TreeNode childNode5 = new TreeNode(currentRowLevel50);
                                                                                                                //TreeNode childNode5 = new TreeNode(currentRowLevel50_Num);

                                                                                                                if (lastRowLevel50 != currentRowLevel50)
                                                                                                                {
                                                                                                                    if (currentRowLevel40_2 == currentRowLevel40)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                                    {
                                                                                                                        childNode4.Nodes.Add(childNode5);



                                                                                                                    }
                                                                                                                    //

                                                                                                                }
                                                                                                            }
                                                                                                        }
                                                                                                    }


                                                                                                }
                                                                                                //

                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                            //

                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                    }
                                                }
                                            }
                                        }

                                        if (i >= table_RowIndex2 && i < table_RowIndex3)
                                        {

                                            for (int j = table_RowIndex2; j < table_RowIndex3; j++)
                                            {

                                                if (j > table_RowIndex2)
                                                {


                                                    string currentRowLevel10 = table.Rows[j]["Group1"].ToString();
                                                    string lastRowLevel10 = table.Rows[table_RowIndex2]["Group1"].ToString();

                                                    string currentRowLevel20 = table.Rows[j]["Group2"].ToString();
                                                    string lastRowLevel20 = table.Rows[j - 1]["Group2"].ToString();

                                                    string currentRowLevel20_Num = table.Rows[j]["Number"].ToString();

                                                    if (lastRowLevel20 != currentRowLevel20)
                                                    {

                                                        TreeNode childNode = new TreeNode(currentRowLevel20_Num + " - " + currentRowLevel20);
                                                        //TreeNode childNode = new TreeNode(currentRowLevel20);
                                                        //TreeNode childNode = new TreeNode(currentRowLevel20_Num);

                                                        if (lastRowLevel10 == currentRowLevel10)
                                                        {
                                                            lastNode.Nodes.Add(childNode);

                                                            //for (int jj = table_RowIndex1; jj < table_RowIndex2; jj += 2)
                                                            for (int jj = table_RowIndex2; jj < table_RowIndex3; jj++)
                                                            {
                                                                if (jj > table_RowIndex2) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                {
                                                                    string currentRowLevel20_2 = table.Rows[jj]["Group2"].ToString();
                                                                    string lastRowLevel20_2 = table.Rows[table_RowIndex0]["Group2"].ToString();

                                                                    string currentRowLevel30 = table.Rows[jj]["Group3"].ToString();
                                                                    string lastRowLevel30 = table.Rows[jj - 1]["Group3"].ToString();

                                                                    string currentRowLevel30_Num = table.Rows[jj]["Number"].ToString();

                                                                    if (currentRowLevel30 != "")
                                                                    {
                                                                        TreeNode childNode3 = new TreeNode(currentRowLevel30_Num + " - " + currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30_Num);


                                                                        if (lastRowLevel30 != currentRowLevel30)
                                                                        {
                                                                            if (currentRowLevel20_2 == currentRowLevel20)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                            {

                                                                                childNode.Nodes.Add(childNode3);


                                                                                for (int jjj = table_RowIndex2; jjj < table_RowIndex3; jjj++)
                                                                                {
                                                                                    if (jjj > table_RowIndex2) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                    {
                                                                                        string currentRowLevel30_2 = table.Rows[jjj]["Group3"].ToString();
                                                                                        string lastRowLevel30_2 = table.Rows[table_RowIndex0]["Group3"].ToString();

                                                                                        string currentRowLevel40 = table.Rows[jjj]["Group4"].ToString();
                                                                                        string lastRowLevel40 = table.Rows[jjj - 1]["Group4"].ToString();

                                                                                        string currentRowLevel40_Num = table.Rows[jjj]["Number"].ToString();

                                                                                        if (currentRowLevel40 != "")
                                                                                        {
                                                                                            TreeNode childNode4 = new TreeNode(currentRowLevel40_Num + " - " + currentRowLevel40);
                                                                                            //TreeNode childNode4 = new TreeNode(currentRowLevel40);
                                                                                            //TreeNode childNode4 = new TreeNode(currentRowLevel40_Num);

                                                                                            if (lastRowLevel40 != currentRowLevel40)
                                                                                            {
                                                                                                if (currentRowLevel30_2 == currentRowLevel30)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                {
                                                                                                    childNode3.Nodes.Add(childNode4);



                                                                                                    //for (int jjjj = table_RowIndex2; jjjj < table_RowIndex3; jjjj++)
                                                                                                    //{
                                                                                                    //    if (jjjj > table_RowIndex2) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                    //    {
                                                                                                    //        string currentRowLevel40_2 = table.Rows[jjjj]["Group4"].ToString();
                                                                                                    //        string lastRowLevel40_2 = table.Rows[table_RowIndex0]["Group4"].ToString();

                                                                                                    //        string currentRowLevel50 = table.Rows[jjjj]["Group5"].ToString();
                                                                                                    //        string lastRowLevel50 = table.Rows[jjjj - 1]["Group5"].ToString();

                                                                                                    //        string currentRowLevel50_Num = table.Rows[jjjj]["Number"].ToString();

                                                                                                    //        if (currentRowLevel50 != "")
                                                                                                    //        {
                                                                                                    //            TreeNode childNode5 = new TreeNode(currentRowLevel50_Num + " - " + currentRowLevel50);
                                                                                                    //            //TreeNode childNode5 = new TreeNode(currentRowLevel50);
                                                                                                    //            //TreeNode childNode5 = new TreeNode(currentRowLevel50_Num);

                                                                                                    //            //if (lastRowLevel50 != currentRowLevel50)
                                                                                                    //            //{
                                                                                                    //            if (currentRowLevel40_2 == currentRowLevel40)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                    //            {
                                                                                                    //                childNode4.Nodes.Add(childNode5);



                                                                                                    //            }
                                                                                                    //            //

                                                                                                    //            //}
                                                                                                    //        }
                                                                                                    //    }
                                                                                                    //}


                                                                                                }


                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                            //

                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                    }
                                                }
                                            }
                                        }

                                        if (i >= table_RowIndex3 && i < table_RowIndex4)
                                        {

                                            for (int j = table_RowIndex3; j < table_RowIndex4; j++)
                                            {

                                                if (j > table_RowIndex3)
                                                {

                                                    string currentRowLevel10 = table.Rows[j]["Group1"].ToString();
                                                    string lastRowLevel10 = table.Rows[table_RowIndex3]["Group1"].ToString();

                                                    string currentRowLevel20 = table.Rows[j]["Group2"].ToString();
                                                    string lastRowLevel20 = table.Rows[j - 1]["Group2"].ToString();

                                                    string currentRowLevel20_Num = table.Rows[j]["Number"].ToString();


                                                    if (lastRowLevel20 != currentRowLevel20)
                                                    {

                                                        TreeNode childNode = new TreeNode(currentRowLevel20_Num + " - " + currentRowLevel20);
                                                        //TreeNode childNode = new TreeNode(currentRowLevel20);
                                                        //TreeNode childNode = new TreeNode(currentRowLevel20_Num);


                                                        if (lastRowLevel10 == currentRowLevel10)
                                                        {
                                                            lastNode.Nodes.Add(childNode);

                                                            //for (int jj = table_RowIndex1; jj < table_RowIndex2; jj += 2)
                                                            for (int jj = table_RowIndex3; jj < table_RowIndex4; jj++)
                                                            {
                                                                if (jj > table_RowIndex3) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                {
                                                                    string currentRowLevel20_2 = table.Rows[jj]["Group2"].ToString();
                                                                    string lastRowLevel20_2 = table.Rows[table_RowIndex0]["Group2"].ToString();

                                                                    string currentRowLevel30 = table.Rows[jj]["Group3"].ToString();
                                                                    string lastRowLevel30 = table.Rows[jj - 1]["Group3"].ToString();

                                                                    string currentRowLevel30_Num = table.Rows[jj]["Number"].ToString();

                                                                    if (currentRowLevel30 != "")
                                                                    {
                                                                        TreeNode childNode3 = new TreeNode(currentRowLevel30_Num + " - " + currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30_Num);


                                                                        if (lastRowLevel30 != currentRowLevel30)
                                                                        {
                                                                            if (currentRowLevel20_2 == currentRowLevel20)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                            {

                                                                                childNode.Nodes.Add(childNode3);


                                                                                for (int jjj = table_RowIndex3; jjj < table_RowIndex4; jjj++)
                                                                                {
                                                                                    if (jjj > table_RowIndex3) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                    {
                                                                                        string currentRowLevel30_2 = table.Rows[jjj]["Group3"].ToString();
                                                                                        string lastRowLevel30_2 = table.Rows[table_RowIndex0]["Group3"].ToString();

                                                                                        string currentRowLevel40 = table.Rows[jjj]["Group4"].ToString();
                                                                                        string lastRowLevel40 = table.Rows[jjj - 1]["Group4"].ToString();

                                                                                        string currentRowLevel40_Num = table.Rows[jjj]["Number"].ToString();

                                                                                        if (currentRowLevel40 != "")
                                                                                        {
                                                                                            TreeNode childNode4 = new TreeNode(currentRowLevel40_Num + " - " + currentRowLevel40);
                                                                                            // TreeNode childNode4 = new TreeNode( currentRowLevel40);
                                                                                            //TreeNode childNode4 = new TreeNode(currentRowLevel40_Num);

                                                                                            if (lastRowLevel40 != currentRowLevel40)
                                                                                            {
                                                                                                if (currentRowLevel30_2 == currentRowLevel30)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                {
                                                                                                    childNode3.Nodes.Add(childNode4);



                                                                                                    for (int jjjj = table_RowIndex3; jjjj < table_RowIndex4; jjjj++)
                                                                                                    {
                                                                                                        if (jjjj > table_RowIndex3) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                        {
                                                                                                            string currentRowLevel40_2 = table.Rows[jjjj]["Group4"].ToString();
                                                                                                            string lastRowLevel40_2 = table.Rows[table_RowIndex0]["Group4"].ToString();

                                                                                                            string currentRowLevel50 = table.Rows[jjjj]["Group5"].ToString();
                                                                                                            string lastRowLevel50 = table.Rows[jjjj - 1]["Group5"].ToString();

                                                                                                            string currentRowLevel50_Num = table.Rows[jjjj]["Number"].ToString();

                                                                                                            if (currentRowLevel50 != "")
                                                                                                            {
                                                                                                                TreeNode childNode5 = new TreeNode(currentRowLevel50_Num + " - " + currentRowLevel50);
                                                                                                                //TreeNode childNode5 = new TreeNode(currentRowLevel50);
                                                                                                                //TreeNode childNode5 = new TreeNode(currentRowLevel50_Num);

                                                                                                                if (lastRowLevel50 != currentRowLevel50)
                                                                                                                {
                                                                                                                    if (currentRowLevel40_2 == currentRowLevel40)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                                    {
                                                                                                                        childNode4.Nodes.Add(childNode5);



                                                                                                                    }
                                                                                                                    //

                                                                                                                }
                                                                                                            }
                                                                                                        }
                                                                                                    }


                                                                                                }
                                                                                                //

                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }

                                                                            }
                                                                            //

                                                                        }
                                                                    }
                                                                }
                                                            }

                                                        }

                                                    }
                                                }
                                            }
                                        }


                                        if (i >= table_RowIndex4 && i < table_RowIndex5)
                                        {

                                            for (int j = table_RowIndex4; j < table_RowIndex5; j++)
                                            {

                                                if (j > table_RowIndex4)
                                                {


                                                    string currentRowLevel10 = table.Rows[j]["Group1"].ToString();
                                                    string lastRowLevel10 = table.Rows[table_RowIndex4]["Group1"].ToString();

                                                    string currentRowLevel20 = table.Rows[j]["Group2"].ToString();
                                                    string lastRowLevel20 = table.Rows[j - 1]["Group2"].ToString();

                                                    string currentRowLevel20_Num = table.Rows[j]["Number"].ToString();


                                                    if (lastRowLevel20 != currentRowLevel20)
                                                    {

                                                        TreeNode childNode = new TreeNode(currentRowLevel20_Num + " - " + currentRowLevel20);
                                                        //TreeNode childNode = new TreeNode(currentRowLevel20);
                                                        //TreeNode childNode = new TreeNode(currentRowLevel20_Num);

                                                        if (lastRowLevel10 == currentRowLevel10)
                                                        {
                                                            lastNode.Nodes.Add(childNode);

                                                            //for (int jj = table_RowIndex1; jj < table_RowIndex2; jj += 2)
                                                            for (int jj = table_RowIndex4; jj < table_RowIndex5; jj++)
                                                            {
                                                                if (jj > table_RowIndex4) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                {
                                                                    string currentRowLevel20_2 = table.Rows[jj]["Group2"].ToString();
                                                                    string lastRowLevel20_2 = table.Rows[table_RowIndex0]["Group2"].ToString();

                                                                    string currentRowLevel30 = table.Rows[jj]["Group3"].ToString();
                                                                    string lastRowLevel30 = table.Rows[jj - 1]["Group3"].ToString();

                                                                    string currentRowLevel30_Num = table.Rows[jj]["Number"].ToString();


                                                                    if (currentRowLevel30 != "")
                                                                    {
                                                                        TreeNode childNode3 = new TreeNode(currentRowLevel30_Num + " - " + currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30_Num);

                                                                        if (lastRowLevel30 != currentRowLevel30)
                                                                        {
                                                                            if (currentRowLevel20_2 == currentRowLevel20)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                            {

                                                                                childNode.Nodes.Add(childNode3);

                                                                                for (int jjj = table_RowIndex4; jjj < table_RowIndex5; jjj++)
                                                                                {
                                                                                    if (jjj > table_RowIndex4) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                    {
                                                                                        string currentRowLevel30_2 = table.Rows[jjj]["Group3"].ToString();
                                                                                        string lastRowLevel30_2 = table.Rows[table_RowIndex0]["Group3"].ToString();

                                                                                        string currentRowLevel40 = table.Rows[jjj]["Group4"].ToString();
                                                                                        string lastRowLevel40 = table.Rows[jjj - 1]["Group4"].ToString();

                                                                                        string currentRowLevel40_Num = table.Rows[jjj]["Number"].ToString();

                                                                                        if (currentRowLevel40 != "")
                                                                                        {
                                                                                            TreeNode childNode4 = new TreeNode(currentRowLevel40_Num + " - " + currentRowLevel40);
                                                                                            //TreeNode childNode4 = new TreeNode( currentRowLevel40);
                                                                                            //TreeNode childNode4 = new TreeNode(currentRowLevel40_Num);


                                                                                            if (lastRowLevel40 != currentRowLevel40)
                                                                                            {
                                                                                                if (currentRowLevel30_2 == currentRowLevel30)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                {
                                                                                                    childNode3.Nodes.Add(childNode4);



                                                                                                    for (int jjjj = table_RowIndex4; jjjj < table_RowIndex5; jjjj++)
                                                                                                    {
                                                                                                        if (jjjj > table_RowIndex4)  //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                        {
                                                                                                            string currentRowLevel40_2 = table.Rows[jjjj]["Group4"].ToString();
                                                                                                            string lastRowLevel40_2 = table.Rows[table_RowIndex0]["Group4"].ToString();

                                                                                                            string currentRowLevel50 = table.Rows[jjjj]["Group5"].ToString();
                                                                                                            string lastRowLevel50 = table.Rows[jjjj - 1]["Group5"].ToString();

                                                                                                            string currentRowLevel50_Num = table.Rows[jjjj]["Number"].ToString();

                                                                                                            if (currentRowLevel50 != "")
                                                                                                            {
                                                                                                                TreeNode childNode5 = new TreeNode(currentRowLevel50_Num + " - " + currentRowLevel50);
                                                                                                                //TreeNode childNode5 = new TreeNode(currentRowLevel50);
                                                                                                                //TreeNode childNode5 = new TreeNode(currentRowLevel50_Num);

                                                                                                                if (lastRowLevel50 != currentRowLevel50)
                                                                                                                {
                                                                                                                    if (currentRowLevel40_2 == currentRowLevel40)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                                    {
                                                                                                                        childNode4.Nodes.Add(childNode5);



                                                                                                                    }
                                                                                                                    //

                                                                                                                }
                                                                                                            }
                                                                                                        }
                                                                                                    }


                                                                                                }
                                                                                                //

                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }

                                                                            }
                                                                            //

                                                                        }
                                                                    }
                                                                }
                                                            }


                                                        }

                                                    }
                                                }
                                            }
                                        }


                                        if (i >= table_RowIndex5 && i < table_RowIndex6)
                                        {

                                            for (int j = table_RowIndex5; j < table_RowIndex6; j++)
                                            {

                                                if (j > table_RowIndex5)
                                                {


                                                    string currentRowLevel10 = table.Rows[j]["Group1"].ToString();
                                                    string lastRowLevel10 = table.Rows[table_RowIndex5]["Group1"].ToString();

                                                    string currentRowLevel20 = table.Rows[j]["Group2"].ToString();
                                                    string lastRowLevel20 = table.Rows[j - 1]["Group2"].ToString();

                                                    string currentRowLevel20_Num = table.Rows[j]["Number"].ToString();

                                                    if (lastRowLevel20 != currentRowLevel20)
                                                    {

                                                        TreeNode childNode = new TreeNode(currentRowLevel20_Num + " - " + currentRowLevel20);
                                                        //TreeNode childNode = new TreeNode(currentRowLevel20);
                                                        //TreeNode childNode = new TreeNode(currentRowLevel20_Num);

                                                        if (lastRowLevel10 == currentRowLevel10)
                                                        {
                                                            lastNode.Nodes.Add(childNode);

                                                            //for (int jj = table_RowIndex1; jj < table_RowIndex2; jj += 2)
                                                            for (int jj = table_RowIndex5; jj < table_RowIndex6; jj++)
                                                            {
                                                                if (jj > table_RowIndex5) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                {
                                                                    string currentRowLevel20_2 = table.Rows[jj]["Group2"].ToString();
                                                                    string lastRowLevel20_2 = table.Rows[table_RowIndex0]["Group2"].ToString();

                                                                    string currentRowLevel30 = table.Rows[jj]["Group3"].ToString();
                                                                    string lastRowLevel30 = table.Rows[jj - 1]["Group3"].ToString();

                                                                    string currentRowLevel30_Num = table.Rows[jj]["Number"].ToString();

                                                                    if (currentRowLevel30 != "")
                                                                    {
                                                                        TreeNode childNode3 = new TreeNode(currentRowLevel30_Num + " - " + currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30_Num);


                                                                        if (lastRowLevel30 != currentRowLevel30)
                                                                        {
                                                                            if (currentRowLevel20_2 == currentRowLevel20)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                            {

                                                                                childNode.Nodes.Add(childNode3);


                                                                                for (int jjj = table_RowIndex5; jjj < table_RowIndex6; jjj++)
                                                                                {
                                                                                    if (jjj > table_RowIndex5) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                    {
                                                                                        string currentRowLevel30_2 = table.Rows[jjj]["Group3"].ToString();
                                                                                        string lastRowLevel30_2 = table.Rows[table_RowIndex0]["Group3"].ToString();

                                                                                        string currentRowLevel40 = table.Rows[jjj]["Group4"].ToString();
                                                                                        string lastRowLevel40 = table.Rows[jjj - 1]["Group4"].ToString();

                                                                                        string currentRowLevel40_Num = table.Rows[jjj]["Number"].ToString();

                                                                                        if (currentRowLevel40 != "")
                                                                                        {
                                                                                            TreeNode childNode4 = new TreeNode(currentRowLevel40_Num + " - " + currentRowLevel40);
                                                                                            //TreeNode childNode4 = new TreeNode( currentRowLevel40);
                                                                                            //TreeNode childNode4 = new TreeNode(currentRowLevel40_Num);


                                                                                            if (lastRowLevel40 != currentRowLevel40)
                                                                                            {
                                                                                                if (currentRowLevel30_2 == currentRowLevel30)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                {
                                                                                                    childNode3.Nodes.Add(childNode4);



                                                                                                    for (int jjjj = table_RowIndex5; jjjj < table_RowIndex6; jjjj++)
                                                                                                    {
                                                                                                        if (jjjj > table_RowIndex5) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                        {
                                                                                                            string currentRowLevel40_2 = table.Rows[jjjj]["Group4"].ToString();
                                                                                                            string lastRowLevel40_2 = table.Rows[table_RowIndex0]["Group4"].ToString();

                                                                                                            string currentRowLevel50 = table.Rows[jjjj]["Group5"].ToString();
                                                                                                            string lastRowLevel50 = table.Rows[jjjj - 1]["Group5"].ToString();

                                                                                                            string currentRowLevel50_Num = table.Rows[jjjj]["Number"].ToString();

                                                                                                            if (currentRowLevel50 != "")
                                                                                                            {
                                                                                                                TreeNode childNode5 = new TreeNode(currentRowLevel50_Num + " - " + currentRowLevel50);
                                                                                                                //TreeNode childNode5 = new TreeNode(currentRowLevel50);
                                                                                                                //TreeNode childNode5 = new TreeNode(currentRowLevel50_Num);

                                                                                                                if (lastRowLevel50 != currentRowLevel50)
                                                                                                                {
                                                                                                                    if (currentRowLevel40_2 == currentRowLevel40)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                                    {
                                                                                                                        childNode4.Nodes.Add(childNode5);



                                                                                                                    }
                                                                                                                    //

                                                                                                                }
                                                                                                            }
                                                                                                        }
                                                                                                    }


                                                                                                }
                                                                                                //

                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                            //

                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                    }
                                                }
                                            }
                                        }

                                        if (i >= table_RowIndex6 && i < table_RowIndex7)
                                        {

                                            for (int j = table_RowIndex6; j < table_RowIndex7; j++)
                                            {

                                                if (j > table_RowIndex6)
                                                {


                                                    string currentRowLevel10 = table.Rows[j]["Group1"].ToString();
                                                    string lastRowLevel10 = table.Rows[table_RowIndex6]["Group1"].ToString();

                                                    string currentRowLevel20 = table.Rows[j]["Group2"].ToString();
                                                    string lastRowLevel20 = table.Rows[j - 1]["Group2"].ToString();

                                                    string currentRowLevel20_Num = table.Rows[j]["Number"].ToString();

                                                    if (lastRowLevel20 != currentRowLevel20)
                                                    {

                                                        TreeNode childNode = new TreeNode(currentRowLevel20_Num + " - " + currentRowLevel20);
                                                        // TreeNode childNode = new TreeNode(currentRowLevel20);
                                                        //TreeNode childNode = new TreeNode(currentRowLevel20_Num);

                                                        if (lastRowLevel10 == currentRowLevel10)
                                                        {
                                                            lastNode.Nodes.Add(childNode);

                                                            //for (int jj = table_RowIndex1; jj < table_RowIndex2; jj += 2)
                                                            for (int jj = table_RowIndex6; jj < table_RowIndex7; jj++)
                                                            {
                                                                if (jj > table_RowIndex6) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                {
                                                                    string currentRowLevel20_2 = table.Rows[jj]["Group2"].ToString();
                                                                    string lastRowLevel20_2 = table.Rows[table_RowIndex0]["Group2"].ToString();

                                                                    string currentRowLevel30 = table.Rows[jj]["Group3"].ToString();
                                                                    string lastRowLevel30 = table.Rows[jj - 1]["Group3"].ToString();

                                                                    string currentRowLevel30_Num = table.Rows[jj]["Number"].ToString();

                                                                    if (currentRowLevel30 != "")
                                                                    {
                                                                        TreeNode childNode3 = new TreeNode(currentRowLevel30_Num + " - " + currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30_Num);


                                                                        if (lastRowLevel30 != currentRowLevel30)
                                                                        {
                                                                            if (currentRowLevel20_2 == currentRowLevel20)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                            {

                                                                                childNode.Nodes.Add(childNode3);


                                                                                for (int jjj = table_RowIndex6; jjj < table_RowIndex7; jjj++)
                                                                                {
                                                                                    if (jjj > table_RowIndex6) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                    {
                                                                                        string currentRowLevel30_2 = table.Rows[jjj]["Group3"].ToString();
                                                                                        string lastRowLevel30_2 = table.Rows[table_RowIndex0]["Group3"].ToString();

                                                                                        string currentRowLevel40 = table.Rows[jjj]["Group4"].ToString();
                                                                                        string lastRowLevel40 = table.Rows[jjj - 1]["Group4"].ToString();

                                                                                        string currentRowLevel40_Num = table.Rows[jjj]["Number"].ToString();

                                                                                        if (currentRowLevel40 != "")
                                                                                        {
                                                                                            TreeNode childNode4 = new TreeNode(currentRowLevel40_Num + " - " + currentRowLevel40);
                                                                                            //TreeNode childNode4 = new TreeNode( currentRowLevel40);
                                                                                            //TreeNode childNode4 = new TreeNode(currentRowLevel40_Num);


                                                                                            if (lastRowLevel40 != currentRowLevel40)
                                                                                            {
                                                                                                if (currentRowLevel30_2 == currentRowLevel30)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                {
                                                                                                    childNode3.Nodes.Add(childNode4);



                                                                                                    for (int jjjj = table_RowIndex6; jjjj < table_RowIndex7; jjjj++)
                                                                                                    {
                                                                                                        if (jjjj > table_RowIndex6) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                        {
                                                                                                            string currentRowLevel40_2 = table.Rows[jjjj]["Group4"].ToString();
                                                                                                            string lastRowLevel40_2 = table.Rows[table_RowIndex0]["Group4"].ToString();

                                                                                                            string currentRowLevel50 = table.Rows[jjjj]["Group5"].ToString();
                                                                                                            string lastRowLevel50 = table.Rows[jjjj - 1]["Group5"].ToString();

                                                                                                            string currentRowLevel50_Num = table.Rows[jjjj]["Number"].ToString();

                                                                                                            if (currentRowLevel50 != "")
                                                                                                            {
                                                                                                                TreeNode childNode5 = new TreeNode(currentRowLevel50_Num + " - " + currentRowLevel50);
                                                                                                                //TreeNode childNode5 = new TreeNode(currentRowLevel50);
                                                                                                                //TreeNode childNode5 = new TreeNode(currentRowLevel50_Num);

                                                                                                                if (lastRowLevel50 != currentRowLevel50)
                                                                                                                {
                                                                                                                    if (currentRowLevel40_2 == currentRowLevel40)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                                    {
                                                                                                                        childNode4.Nodes.Add(childNode5);



                                                                                                                    }
                                                                                                                    //

                                                                                                                }
                                                                                                            }
                                                                                                        }
                                                                                                    }


                                                                                                }
                                                                                                //

                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }

                                                                            }
                                                                            //

                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                    }
                                                }
                                            }
                                        }

                                        if (i >= table_RowIndex7 && i < table_RowIndex8)
                                        {

                                            for (int j = table_RowIndex7; j < table_RowIndex8; j++)
                                            {

                                                if (j > table_RowIndex7)
                                                {


                                                    string currentRowLevel10 = table.Rows[j]["Group1"].ToString();
                                                    string lastRowLevel10 = table.Rows[table_RowIndex7]["Group1"].ToString();

                                                    string currentRowLevel20 = table.Rows[j]["Group2"].ToString();
                                                    string lastRowLevel20 = table.Rows[j - 1]["Group2"].ToString();

                                                    string currentRowLevel20_Num = table.Rows[j]["Number"].ToString();

                                                    if (lastRowLevel20 != currentRowLevel20)
                                                    {

                                                        TreeNode childNode = new TreeNode(currentRowLevel20_Num + " - " + currentRowLevel20);
                                                        //TreeNode childNode = new TreeNode(currentRowLevel20);
                                                        //TreeNode childNode = new TreeNode(currentRowLevel20_Num);

                                                        if (lastRowLevel10 == currentRowLevel10)
                                                        {
                                                            lastNode.Nodes.Add(childNode);


                                                            //for (int jj = table_RowIndex1; jj < table_RowIndex2; jj += 2)
                                                            for (int jj = table_RowIndex7; jj < table_RowIndex8; jj++)
                                                            {
                                                                if (jj > table_RowIndex7) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                {
                                                                    string currentRowLevel20_2 = table.Rows[jj]["Group2"].ToString();
                                                                    string lastRowLevel20_2 = table.Rows[table_RowIndex0]["Group2"].ToString();

                                                                    string currentRowLevel30 = table.Rows[jj]["Group3"].ToString();
                                                                    string lastRowLevel30 = table.Rows[jj - 1]["Group3"].ToString();

                                                                    string currentRowLevel30_Num = table.Rows[jj]["Number"].ToString();

                                                                    if (currentRowLevel30 != "")
                                                                    {
                                                                        TreeNode childNode3 = new TreeNode(currentRowLevel30_Num + " - " + currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30_Num);


                                                                        if (lastRowLevel30 != currentRowLevel30)
                                                                        {
                                                                            if (currentRowLevel20_2 == currentRowLevel20)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                            {

                                                                                childNode.Nodes.Add(childNode3);

                                                                                for (int jjj = table_RowIndex7; jjj < table_RowIndex8; jjj++)
                                                                                {
                                                                                    if (jjj > table_RowIndex7) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                    {
                                                                                        string currentRowLevel30_2 = table.Rows[jjj]["Group3"].ToString();
                                                                                        string lastRowLevel30_2 = table.Rows[table_RowIndex0]["Group3"].ToString();

                                                                                        string currentRowLevel40 = table.Rows[jjj]["Group4"].ToString();
                                                                                        string lastRowLevel40 = table.Rows[jjj - 1]["Group4"].ToString();

                                                                                        string currentRowLevel40_Num = table.Rows[jjj]["Number"].ToString();

                                                                                        if (currentRowLevel40 != "")
                                                                                        {
                                                                                            TreeNode childNode4 = new TreeNode(currentRowLevel40_Num + " - " + currentRowLevel40);
                                                                                            //TreeNode childNode4 = new TreeNode( currentRowLevel40);
                                                                                            //TreeNode childNode4 = new TreeNode(currentRowLevel40_Num);


                                                                                            if (lastRowLevel40 != currentRowLevel40)
                                                                                            {
                                                                                                if (currentRowLevel30_2 == currentRowLevel30)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                {
                                                                                                    childNode3.Nodes.Add(childNode4);


                                                                                                    for (int jjjj = table_RowIndex7; jjjj < table_RowIndex8; jjjj++)
                                                                                                    {
                                                                                                        if (jjjj > table_RowIndex7)  //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                        {
                                                                                                            string currentRowLevel40_2 = table.Rows[jjjj]["Group4"].ToString();
                                                                                                            string lastRowLevel40_2 = table.Rows[table_RowIndex0]["Group4"].ToString();

                                                                                                            string currentRowLevel50 = table.Rows[jjjj]["Group5"].ToString();
                                                                                                            string lastRowLevel50 = table.Rows[jjjj - 1]["Group5"].ToString();

                                                                                                            string currentRowLevel50_Num = table.Rows[jjjj]["Number"].ToString();

                                                                                                            if (currentRowLevel50 != "")
                                                                                                            {
                                                                                                                TreeNode childNode5 = new TreeNode(currentRowLevel50_Num + " - " + currentRowLevel50);
                                                                                                                //TreeNode childNode5 = new TreeNode(currentRowLevel50);
                                                                                                                //TreeNode childNode5 = new TreeNode(currentRowLevel50_Num);

                                                                                                                if (lastRowLevel50 != currentRowLevel50)
                                                                                                                {
                                                                                                                    if (currentRowLevel40_2 == currentRowLevel40)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                                    {
                                                                                                                        childNode4.Nodes.Add(childNode5);



                                                                                                                    }
                                                                                                                    //

                                                                                                                }
                                                                                                            }
                                                                                                        }
                                                                                                    }



                                                                                                }
                                                                                                //

                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }

                                                                            }
                                                                            //

                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                    }
                                                }
                                            }
                                        }


                                        if (i >= table_RowIndex8 && i <= table_RowIndexSum)
                                        {

                                            for (int j = table_RowIndex8; j < table_RowIndexSum + 1; j++)
                                            {

                                                if (j > table_RowIndex8)
                                                {


                                                    string currentRowLevel10 = table.Rows[j]["Group1"].ToString();
                                                    string lastRowLevel10 = table.Rows[table_RowIndex8]["Group1"].ToString();

                                                    string currentRowLevel20 = table.Rows[j]["Group2"].ToString();
                                                    string lastRowLevel20 = table.Rows[j - 1]["Group2"].ToString();

                                                    string currentRowLevel20_Num = table.Rows[j]["Number"].ToString();

                                                    if (lastRowLevel20 != currentRowLevel20)
                                                    {

                                                        TreeNode childNode = new TreeNode(currentRowLevel20_Num + " - " + currentRowLevel20);
                                                        //TreeNode childNode = new TreeNode(currentRowLevel20);
                                                        //TreeNode childNode = new TreeNode(currentRowLevel20_Num);

                                                        if (lastRowLevel10 == currentRowLevel10)
                                                        {
                                                            lastNode.Nodes.Add(childNode);

                                                            //for (int jj = table_RowIndex1; jj < table_RowIndex2; jj += 2)
                                                            for (int jj = table_RowIndex8; jj < table_RowIndexSum + 1; jj++)
                                                            {
                                                                if (jj > table_RowIndex8) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                {
                                                                    string currentRowLevel20_2 = table.Rows[jj]["Group2"].ToString();
                                                                    string lastRowLevel20_2 = table.Rows[table_RowIndex0]["Group2"].ToString();

                                                                    string currentRowLevel30 = table.Rows[jj]["Group3"].ToString();
                                                                    string lastRowLevel30 = table.Rows[jj - 1]["Group3"].ToString();

                                                                    string currentRowLevel30_Num = table.Rows[jj]["Number"].ToString();

                                                                    if (currentRowLevel30 != "")
                                                                    {
                                                                        TreeNode childNode3 = new TreeNode(currentRowLevel30_Num + " - " + currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30);
                                                                        //TreeNode childNode3 = new TreeNode(currentRowLevel30_Num);


                                                                        if (lastRowLevel30 != currentRowLevel30)
                                                                        {
                                                                            if (currentRowLevel20_2 == currentRowLevel20)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                            {

                                                                                childNode.Nodes.Add(childNode3);

                                                                                for (int jjj = table_RowIndex8; jjj < table_RowIndexSum + 1; jjj++)
                                                                                {
                                                                                    if (jjj > table_RowIndex8) //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                    {
                                                                                        string currentRowLevel30_2 = table.Rows[jjj]["Group3"].ToString();
                                                                                        string lastRowLevel30_2 = table.Rows[table_RowIndex0]["Group3"].ToString();

                                                                                        string currentRowLevel40 = table.Rows[jjj]["Group4"].ToString();
                                                                                        string lastRowLevel40 = table.Rows[jjj - 1]["Group4"].ToString();

                                                                                        string currentRowLevel40_Num = table.Rows[jjj]["Number"].ToString();

                                                                                        if (currentRowLevel40 != "")
                                                                                        {
                                                                                            TreeNode childNode4 = new TreeNode(currentRowLevel40_Num + " - " + currentRowLevel40);
                                                                                            //TreeNode childNode4 = new TreeNode( currentRowLevel40);
                                                                                            //TreeNode childNode4 = new TreeNode(currentRowLevel40_Num);


                                                                                            if (lastRowLevel40 != currentRowLevel40)
                                                                                            {
                                                                                                if (currentRowLevel30_2 == currentRowLevel30)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                {
                                                                                                    childNode3.Nodes.Add(childNode4);


                                                                                                    for (int jjjj = table_RowIndex8; jjjj < table_RowIndexSum + 1; jjjj++)
                                                                                                    {
                                                                                                        if (jjjj > table_RowIndex8)  //  &&  table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                        {
                                                                                                            string currentRowLevel40_2 = table.Rows[jjjj]["Group4"].ToString();
                                                                                                            string lastRowLevel40_2 = table.Rows[table_RowIndex0]["Group4"].ToString();

                                                                                                            string currentRowLevel50 = table.Rows[jjjj]["Group5"].ToString();
                                                                                                            string lastRowLevel50 = table.Rows[jjjj - 1]["Group5"].ToString();

                                                                                                            string currentRowLevel50_Num = table.Rows[jjjj]["Number"].ToString();

                                                                                                            if (currentRowLevel50 != "")
                                                                                                            {
                                                                                                                TreeNode childNode5 = new TreeNode(currentRowLevel50_Num + " - " + currentRowLevel50);
                                                                                                                //TreeNode childNode5 = new TreeNode(currentRowLevel50);
                                                                                                                //TreeNode childNode5 = new TreeNode(currentRowLevel50_Num);

                                                                                                                if (lastRowLevel50 != currentRowLevel50)
                                                                                                                {
                                                                                                                    if (currentRowLevel40_2 == currentRowLevel40)// && table.Rows[jj]["Group3"] != DBNull.Value)
                                                                                                                    {
                                                                                                                        childNode4.Nodes.Add(childNode5);



                                                                                                                    }
                                                                                                                    //

                                                                                                                }
                                                                                                            }
                                                                                                        }
                                                                                                    }



                                                                                                }
                                                                                                //

                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                            //

                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                    }
                                                }
                                            }
                                        }




                                    }




                                }


                            }

                            else
                            {

                                treeView2.Nodes.Add(newNode);





                            }

                        }


                        lastNode = newNode;
                    }
                }
            }


            ///////////////////////////////////////////////////////


                       

        }


        public void TransferQuantitiesToGifaGridView()
        {

            

            string DescriptionTarget = "";
            string MaterialDescription = "";
            string NetVolume = "";
            
                      

           // var foundAuthors = tbl.Select("Author = '" + searchAuthor + "'");
            for (int i = 0; i < GifaDataGridView.Rows.Count; i++)
            {


                if (GifaDataGridView[15, i].Value.ToString() != "")
                {
                    DescriptionTarget = GifaDataGridView[15, i].Value.ToString();

                   

                    for (int j = 0; j < M_QuantitiesTableValue.Rows.Count; j++)
                    {
                        
                        MaterialDescription = M_QuantitiesTableValue.Rows[j][0].ToString();
                        NetVolume = M_QuantitiesTableValue.Rows[j][1].ToString();

                        if (DescriptionTarget == MaterialDescription)
                            GifaDataGridView[2, i].Value = NetVolume;

                        
                    }
                }
            }

               
        }



        public void CreateMatirixByTranspose()
        {


            DataSet ds = new DataSet();
            //ds.DataSetName = "MyData2";

            ds.Tables.Add(ReportElementsInfo());

            //dataGridView3.DataSource = ds.Tables[0];


            DataTable inputTable = ds.Tables[0];
            // Table shown in Figure 1.1

            DataTable transposedTable = GenerateTransposedTable(inputTable);




            // string expression = " Item = 'Substructure'";
            string Gifa0_Expression = " Number = '0'";
            DataRow[] Gifa0_foundRows = inputTable.Select(Gifa0_Expression); ;
            Gifa0_RowToColIndex = inputTable.Rows.IndexOf(Gifa0_foundRows[0]);

            string Gifa1_Expression = " Number = '1'";
            DataRow[] Gifa1_foundRows = inputTable.Select(Gifa1_Expression); ;
            Gifa1_RowToColIndex = inputTable.Rows.IndexOf(Gifa1_foundRows[0]);


            string Gifa2_Expression = " Number = '2'";
            DataRow[] Gifa2_foundRows = inputTable.Select(Gifa2_Expression); ;
            Gifa2_RowToColIndex = inputTable.Rows.IndexOf(Gifa2_foundRows[0]);


            string Gifa3_Expression = " Number = '3'";
            DataRow[] Gifa3_foundRows = inputTable.Select(Gifa3_Expression); ;
            Gifa3_RowToColIndex = inputTable.Rows.IndexOf(Gifa3_foundRows[0]);


            string Gifa4_Expression = " Number = '4'";
            DataRow[] Gifa4_foundRows = inputTable.Select(Gifa4_Expression); ;
            Gifa4_RowToColIndex = inputTable.Rows.IndexOf(Gifa4_foundRows[0]);

            string Gifa5_Expression = " Number = '5'";
            DataRow[] Gifa5_foundRows = inputTable.Select(Gifa5_Expression); ;
            Gifa5_RowToColIndex = inputTable.Rows.IndexOf(Gifa5_foundRows[0]);

            string Gifa6_Expression = " Number = '6'";
            DataRow[] Gifa6_foundRows = inputTable.Select(Gifa6_Expression); ;
            Gifa6_RowToColIndex = inputTable.Rows.IndexOf(Gifa6_foundRows[0]);

            string Gifa7_Expression = " Number = '7'";
            DataRow[] Gifa7_foundRows = inputTable.Select(Gifa7_Expression); ;
            Gifa7_RowToColIndex = inputTable.Rows.IndexOf(Gifa7_foundRows[0]);

            string Gifa8_Expression = " Number = '8'";
            DataRow[] Gifa8_foundRows = inputTable.Select(Gifa8_Expression); ;
            Gifa8_RowToColIndex = inputTable.Rows.IndexOf(Gifa8_foundRows[0]);

            string GifaSum_Expression = " Number = '8.8.3.1'";
            DataRow[] GifaSum_foundRows = inputTable.Select(GifaSum_Expression); ;
            GifaSum_RowToColIndex = inputTable.Rows.IndexOf(GifaSum_foundRows[0]);
            // GifaSum_RowToColIndex = table_RowIndexSum;


            Matrix_mTable = new DataTable();

            Matrix_mTable = transposedTable.Copy();

            ModifiedTransposedTable = transposedTable.Copy();

            //; //Put your column X number here
            for (int i = 0; i <= ModifiedTransposedTable.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= ModifiedTransposedTable.Columns.Count - 1; j++)
                {

                    string ValueCheck = ModifiedTransposedTable.Rows[i][j].ToString();

                    if (ValueCheck == "")
                    {
                        ModifiedTransposedTable.Rows[i][j] = "0";
                    }

                    else
                    {

                        ModifiedTransposedTable.Rows[i][j] = transposedTable.Rows[i][j].ToString();
                    }

                    //MessageBox.Show(ModifiedTransposedTable.Rows[i][j].ToString());
                }
            }


            /////////////////
           



            //; //Put your column X number here
            if (Matrix_mTable.Rows[0] != null)
            {

                for (int j = 1; j <= Gifa1_RowToColIndex - 1; j++)
                {

                    string ValueCheck = transposedTable.Rows[4][j].ToString();
                    //MessageBox.Show(  "ValueCheck = " + ValueCheck);

                    if (ValueCheck == "")
                    {
                        Matrix_mTable.Rows[0][j] = "0";
                    }

                    else
                    {

                        Matrix_mTable.Rows[0][j] = transposedTable.Rows[4][j].ToString();


                    }

                    // MessageBox.Show( "1; = " + Matrix_mTable.Rows[0][j].ToString());

                }

                for (int j = Gifa1_RowToColIndex; j <= Matrix_mTable.Columns.Count - 1; j++)
                {

                    Matrix_mTable.Rows[0][j] = "0";

                }



            }


            if (Matrix_mTable.Rows[1] != null)
            {

                for (int j = 1; j <= Gifa1_RowToColIndex - 1; j++)
                {

                    Matrix_mTable.Rows[1][j] = "0";

                }

                for (int j = Gifa1_RowToColIndex; j <= Gifa2_RowToColIndex - 1; j++)
                {

                    string ValueCheck = transposedTable.Rows[4][j].ToString();

                    if (ValueCheck == "")
                    {
                        Matrix_mTable.Rows[1][j] = "0";
                    }

                    else
                    {

                        Matrix_mTable.Rows[1][j] = transposedTable.Rows[4][j].ToString();


                    }

                    //MessageBox.Show("1; = " + Matrix_mTable.Rows[1][j].ToString());

                }

                for (int j = Gifa2_RowToColIndex; j <= Matrix_mTable.Columns.Count - 1; j++)
                {

                    Matrix_mTable.Rows[1][j] = "0";

                }


            }


            if (Matrix_mTable.Rows[2] != null)
            {

                for (int j = 1; j <= Gifa2_RowToColIndex - 1; j++)
                {

                    Matrix_mTable.Rows[2][j] = "0";

                }

                for (int j = Gifa2_RowToColIndex; j <= Gifa3_RowToColIndex - 1; j++)
                {

                    string ValueCheck = transposedTable.Rows[4][j].ToString();

                    if (ValueCheck == "")
                    {
                        Matrix_mTable.Rows[2][j] = "0";
                    }

                    else
                    {

                        Matrix_mTable.Rows[2][j] = transposedTable.Rows[4][j].ToString();
                    }


                }


                for (int j = Gifa3_RowToColIndex; j <= Matrix_mTable.Columns.Count - 1; j++)
                {

                    Matrix_mTable.Rows[2][j] = "0";

                }
            }


            if (Matrix_mTable.Rows[3] != null)
            {

                for (int j = 1; j <= Gifa3_RowToColIndex - 1; j++)
                {

                    Matrix_mTable.Rows[3][j] = "0";

                }

                for (int j = Gifa3_RowToColIndex; j <= Gifa4_RowToColIndex - 1; j++)
                {

                    string ValueCheck = transposedTable.Rows[4][j].ToString();

                    if (ValueCheck == "")
                    {
                        Matrix_mTable.Rows[3][j] = "0";
                    }

                    else
                    {

                        Matrix_mTable.Rows[3][j] = transposedTable.Rows[4][j].ToString();
                    }

                }

                for (int j = Gifa4_RowToColIndex; j <= Matrix_mTable.Columns.Count - 1; j++)
                {

                    Matrix_mTable.Rows[3][j] = "0";

                }

            }


            if (Matrix_mTable.Rows[4] != null)
            {

                for (int j = 1; j <= Gifa4_RowToColIndex - 1; j++)
                {

                    Matrix_mTable.Rows[4][j] = "0";

                }

                for (int j = Gifa4_RowToColIndex; j <= Gifa5_RowToColIndex - 1; j++)
                {

                    string ValueCheck = transposedTable.Rows[4][j].ToString();

                    if (ValueCheck == "")
                    {
                        Matrix_mTable.Rows[4][j] = "0";
                    }

                    else
                    {

                        Matrix_mTable.Rows[4][j] = transposedTable.Rows[4][j].ToString();
                    }

                }

                for (int j = Gifa5_RowToColIndex; j <= Matrix_mTable.Columns.Count - 1; j++)
                {

                    Matrix_mTable.Rows[4][j] = "0";

                }

            }


            if (Matrix_mTable.Rows[5] != null)
            {

                for (int j = 1; j <= Gifa5_RowToColIndex - 1; j++)
                {

                    Matrix_mTable.Rows[5][j] = "0";

                }

                for (int j = Gifa5_RowToColIndex; j <= Gifa6_RowToColIndex - 1; j++)
                {

                    string ValueCheck = transposedTable.Rows[4][j].ToString();

                    if (ValueCheck == "")
                    {
                        Matrix_mTable.Rows[5][j] = "0";
                    }

                    else
                    {

                        Matrix_mTable.Rows[5][j] = transposedTable.Rows[4][j].ToString();
                    }

                }

                for (int j = Gifa6_RowToColIndex; j <= Matrix_mTable.Columns.Count - 1; j++)
                {

                    Matrix_mTable.Rows[5][j] = "0";

                }

            }


            if (Matrix_mTable.Rows[6] != null)
            {

                for (int j = 1; j <= Gifa6_RowToColIndex - 1; j++)
                {

                    Matrix_mTable.Rows[6][j] = "0";

                }

                for (int j = Gifa6_RowToColIndex; j <= Gifa7_RowToColIndex - 1; j++)
                {

                    string ValueCheck = transposedTable.Rows[4][j].ToString();

                    if (ValueCheck == "")
                    {
                        Matrix_mTable.Rows[6][j] = "0";
                    }

                    else
                    {

                        Matrix_mTable.Rows[6][j] = transposedTable.Rows[4][j].ToString();
                    }
                }

                for (int j = Gifa7_RowToColIndex; j <= Matrix_mTable.Columns.Count - 1; j++)
                {

                    Matrix_mTable.Rows[6][j] = "0";

                }

            }



            if (Matrix_mTable.Rows[7] != null)
            {

                for (int j = 1; j <= Gifa7_RowToColIndex - 1; j++)
                {


                    Matrix_mTable.Rows[7][j] = "0";

                }

                for (int j = Gifa7_RowToColIndex; j <= Gifa8_RowToColIndex - 1; j++)
                //for (int j = Gifa7_RowToColIndex; j <= Matrix_mTable.Columns.Count - 1; j++)
                {

                    string ValueCheck = transposedTable.Rows[4][j].ToString();

                    if (ValueCheck == "")
                    {
                        Matrix_mTable.Rows[7][j] = "0";
                    }

                    else
                    {

                        Matrix_mTable.Rows[7][j] = transposedTable.Rows[4][j].ToString();



                    }

                    //MessageBox.Show("8; = " + Matrix_mTable.Rows[7][j].ToString());

                }
                for (int j = Gifa8_RowToColIndex; j <= Matrix_mTable.Columns.Count - 1; j++)
                {

                    Matrix_mTable.Rows[7][j] = "0";

                }



            }

            ////////////////
            ////////////////

            if (Matrix_mTable.Rows[8] != null)
            {

                for (int j = 1; j <= Gifa8_RowToColIndex - 1; j++)
                {


                    Matrix_mTable.Rows[8][j] = "0";

                }

                for (int j = Gifa8_RowToColIndex; j <= Matrix_mTable.Columns.Count - 1; j++)
                {

                    string ValueCheck = transposedTable.Rows[4][j].ToString();

                    if (ValueCheck == "")
                    {
                        Matrix_mTable.Rows[8][j] = "0";
                    }

                    else
                    {

                        Matrix_mTable.Rows[8][j] = transposedTable.Rows[4][j].ToString();



                    }


                    //MessageBox.Show("8; = " + Matrix_mTable.Rows[7][j].ToString());

                }



            }




        }


        private DataTable ReportElementsInfo()
        {


            DataTable dt = new DataTable();

            dt.TableName = "MyTable 0";




            foreach (DataGridViewColumn col in GifaDataGridView.Columns)
            {
                dt.Columns.Add(col.HeaderText);
            }

           

            foreach (DataGridViewRow gridRow in GifaDataGridView.Rows)
            {
                if (gridRow.IsNewRow)
                    continue;
                DataRow dtRow = dt.NewRow();

                for (int i1 = 0; i1 < GifaDataGridView.Columns.Count; i1++)
                    dtRow[i1] = (gridRow.Cells[i1].Value == null ? DBNull.Value : gridRow.Cells[i1].Value);
                dt.Rows.Add(dtRow);
            }




            return dt;
        }



      

        private DataTable GenerateTransposedTable(DataTable inputTable)
        {
            DataTable outputTable = new DataTable();

            // Add columns by looping rows

            // Header row's first column is same as in inputTable
            outputTable.Columns.Add(inputTable.Columns[0].ColumnName.ToString());

            // Header row's second column onwards, 'inputTable's first column taken
            foreach (DataRow inRow in inputTable.Rows)
            {
                string newColName = inRow[0].ToString();
                outputTable.Columns.Add(newColName);
            }

            // Add rows by looping columns        
            for (int rCount = 1; rCount <= inputTable.Columns.Count - 1; rCount++)
            {
                DataRow newRow = outputTable.NewRow();

                // First column is inputTable's Header row's second column
                newRow[0] = inputTable.Columns[rCount].ColumnName.ToString();
                for (int cCount = 0; cCount <= inputTable.Rows.Count - 1; cCount++)
                {
                    string colValue = inputTable.Rows[cCount][rCount].ToString();
                    newRow[cCount + 1] = colValue;
                }
                outputTable.Rows.Add(newRow);
            }

            return outputTable;
        }


        // --------------------------------------------------------------------------------------------------------------

        static double[][] MatrixCreate(int rows, int cols)
        {
            // allocates/creates a matrix initialized to all 0.0. assume rows and cols > 0
            // do error checking here
            double[][] result = new double[rows][];
            for (int i = 0; i < rows; ++i)
                result[i] = new double[cols];

            //for (int i = 0; i < rows; ++i)
            //  for (int j = 0; j < cols; ++j)
            //    result[i][j] = 0.0; // explicit initialization needed in some languages

            return result;
        }

        // --------------------------------------------------------------------------------------------------------------

        static double[][] MatrixRandom(int rows, int cols, double minVal, double maxVal, int seed)
        {
            // return a matrix with random values
            Random ran = new Random(seed);
            double[][] result = MatrixCreate(rows, cols);
            for (int i = 0; i < rows; ++i)
                for (int j = 0; j < cols; ++j)
                    result[i][j] = (maxVal - minVal) * ran.NextDouble() + minVal;
            return result;
        }

        // --------------------------------------------------------------------------------------------------------------

        static double[][] MatrixIdentity(int n)
        {
            // return an n x n Identity matrix
            double[][] result = MatrixCreate(n, n);
            for (int i = 0; i < n; ++i)
                result[i][i] = 1.0;

            return result;
        }

        // --------------------------------------------------------------------------------------------------------------

        static string MatrixAsString(double[][] matrix)
        {
            string s = "";
            for (int i = 0; i < matrix.Length; ++i)
            {
                for (int j = 0; j < matrix[i].Length; ++j)
                    s += matrix[i][j].ToString("F3").PadLeft(8) + " ";
                s += Environment.NewLine;
            }
            return s;
        }

        // --------------------------------------------------------------------------------------------------------------

        static bool MatrixAreEqual(double[][] matrixA, double[][] matrixB, double epsilon)
        {
            // true if all values in matrixA == corresponding values in matrixB
            int aRows = matrixA.Length; int aCols = matrixA[0].Length;
            int bRows = matrixB.Length; int bCols = matrixB[0].Length;
            if (aRows != bRows || aCols != bCols)
                throw new Exception("Non-conformable matrices in MatrixAreEqual");

            for (int i = 0; i < aRows; ++i) // each row of A and B
                for (int j = 0; j < aCols; ++j) // each col of A and B
                    //if (matrixA[i][j] != matrixB[i][j])
                    if (Math.Abs(matrixA[i][j] - matrixB[i][j]) > epsilon)
                        return false;
            return true;
        }

        // --------------------------------------------------------------------------------------------------------------

        static double[][] MatrixProduct(double[][] matrixA, double[][] matrixB)
        {
            int aRows = matrixA.Length; int aCols = matrixA[0].Length;
            int bRows = matrixB.Length; int bCols = matrixB[0].Length;
            if (aCols != bRows)
                throw new Exception("Non-conformable matrices in MatrixProduct");

            double[][] result = MatrixCreate(aRows, bCols);

            for (int i = 0; i < aRows; ++i) // each row of A
                for (int j = 0; j < bCols; ++j) // each col of B
                    for (int k = 0; k < aCols; ++k) // could use k < bRows
                        result[i][j] += matrixA[i][k] * matrixB[k][j];

            //Parallel.For(0, aRows, i =>
            //  {
            //    for (int j = 0; j < bCols; ++j) // each col of B
            //      for (int k = 0; k < aCols; ++k) // could use k < bRows
            //        result[i][j] += matrixA[i][k] * matrixB[k][j];
            //  }
            //);

            return result;
        }

        // --------------------------------------------------------------------------------------------------------------

        static double[] MatrixVectorProduct(double[][] matrix, double[] vector)
        {
            // result of multiplying an n x m matrix by a m x 1 column vector (yielding an n x 1 column vector)
            int mRows = matrix.Length; int mCols = matrix[0].Length;
            int vRows = vector.Length;
            if (mCols != vRows)
                throw new Exception("Non-conformable matrix and vector in MatrixVectorProduct");
            double[] result = new double[mRows]; // an n x m matrix times a m x 1 column vector is a n x 1 column vector
            for (int i = 0; i < mRows; ++i)
                for (int j = 0; j < mCols; ++j)
                    result[i] += matrix[i][j] * vector[j];
            return result;
        }

        // --------------------------------------------------------------------------------------------------------------

        static double[][] MatrixDecompose(double[][] matrix, out int[] perm, out int toggle)
        {
            // Doolittle LUP decomposition with partial pivoting.
            // rerturns: result is L (with 1s on diagonal) and U; perm holds row permutations; toggle is +1 or -1 (even or odd)
            int rows = matrix.Length;
            int cols = matrix[0].Length; // assume all rows have the same number of columns so just use row [0].
            if (rows != cols)
                throw new Exception("Attempt to MatrixDecompose a non-square mattrix");

            int n = rows; // convenience

            double[][] result = MatrixDuplicate(matrix); // make a copy of the input matrix

            perm = new int[n]; // set up row permutation result
            for (int i = 0; i < n; ++i) { perm[i] = i; }

            toggle = 1; // toggle tracks row swaps. +1 -> even, -1 -> odd. used by MatrixDeterminant

            for (int j = 0; j < n - 1; ++j) // each column
            {
                double colMax = Math.Abs(result[j][j]); // find largest value in col j
                int pRow = j;
                for (int i = j + 1; i < n; ++i)
                {
                    if (result[i][j] > colMax)
                    {
                        colMax = result[i][j];
                        pRow = i;
                    }
                }

                if (pRow != j) // if largest value not on pivot, swap rows
                {
                    double[] rowPtr = result[pRow];
                    result[pRow] = result[j];
                    result[j] = rowPtr;

                    int tmp = perm[pRow]; // and swap perm info
                    perm[pRow] = perm[j];
                    perm[j] = tmp;

                    toggle = -toggle; // adjust the row-swap toggle
                }

                if (Math.Abs(result[j][j]) < 1.0E-20) // if diagonal after swap is zero . . .
                    return null; // consider a throw

                for (int i = j + 1; i < n; ++i)
                {
                    result[i][j] /= result[j][j];
                    for (int k = j + 1; k < n; ++k)
                    {
                        result[i][k] -= result[i][j] * result[j][k];
                    }
                }
            } // main j column loop

            return result;
        } // MatrixDecompose

        // --------------------------------------------------------------------------------------------------------------

        static double[][] MatrixInverse(double[][] matrix)
        {
            int n = matrix.Length;
            double[][] result = MatrixDuplicate(matrix);

            int[] perm;
            int toggle;
            double[][] lum = MatrixDecompose(matrix, out perm, out toggle);
            if (lum == null)
                throw new Exception("Unable to compute inverse");

            double[] b = new double[n];
            for (int i = 0; i < n; ++i)
            {
                for (int j = 0; j < n; ++j)
                {
                    if (i == perm[j])
                        b[j] = 1.0;
                    else
                        b[j] = 0.0;
                }

                double[] x = HelperSolve(lum, b); // 

                for (int j = 0; j < n; ++j)
                    result[j][i] = x[j];
            }
            return result;
        }

        // --------------------------------------------------------------------------------------------------------------

        static double MatrixDeterminant(double[][] matrix)
        {
            int[] perm;
            int toggle;
            double[][] lum = MatrixDecompose(matrix, out perm, out toggle);
            if (lum == null)
                throw new Exception("Unable to compute MatrixDeterminant");
            double result = toggle;
            for (int i = 0; i < lum.Length; ++i)
                result *= lum[i][i];
            return result;
        }

        // --------------------------------------------------------------------------------------------------------------

        static double[] HelperSolve(double[][] luMatrix, double[] b) // helper
        {
            // before calling this helper, permute b using the perm array from MatrixDecompose that generated luMatrix
            int n = luMatrix.Length;
            double[] x = new double[n];
            b.CopyTo(x, 0);

            for (int i = 1; i < n; ++i)
            {
                double sum = x[i];
                for (int j = 0; j < i; ++j)
                    sum -= luMatrix[i][j] * x[j];
                x[i] = sum;
            }

            x[n - 1] /= luMatrix[n - 1][n - 1];
            for (int i = n - 2; i >= 0; --i)
            {
                double sum = x[i];
                for (int j = i + 1; j < n; ++j)
                    sum -= luMatrix[i][j] * x[j];
                x[i] = sum / luMatrix[i][i];
            }

            return x;
        }

        // --------------------------------------------------------------------------------------------------------------

        static double[] SystemSolve(double[][] A, double[] b)
        {
            // Solve Ax = b
            int n = A.Length;

            // 1. decompose A
            int[] perm;
            int toggle;
            double[][] luMatrix = MatrixDecompose(A, out perm, out toggle);
            if (luMatrix == null)
                return null;

            // 2. permute b according to perm[] into bp
            double[] bp = new double[b.Length];
            for (int i = 0; i < n; ++i)
                bp[i] = b[perm[i]];

            // 3. call helper
            double[] x = HelperSolve(luMatrix, bp);
            return x;
        } // SystemSolve

        // --------------------------------------------------------------------------------------------------------------

        static double[][] MatrixDuplicate(double[][] matrix)
        {
            // allocates/creates a duplicate of a matrix. assumes matrix is not null.
            double[][] result = MatrixCreate(matrix.Length, matrix[0].Length);
            for (int i = 0; i < matrix.Length; ++i) // copy the values
                for (int j = 0; j < matrix[i].Length; ++j)
                    result[i][j] = matrix[i][j];
            return result;
        }

        // --------------------------------------------------------------------------------------------------------------

        static double[][] ExtractLower(double[][] matrix)
        {
            // lower part of a Doolittle decomposition (1.0s on diagonal, 0.0s in upper)
            int rows = matrix.Length; int cols = matrix[0].Length;
            double[][] result = MatrixCreate(rows, cols);
            for (int i = 0; i < rows; ++i)
            {
                for (int j = 0; j < cols; ++j)
                {
                    if (i == j)
                        result[i][j] = 1.0;
                    else if (i > j)
                        result[i][j] = matrix[i][j];
                }
            }
            return result;
        }

        static double[][] ExtractUpper(double[][] matrix)
        {
            // upper part of a Doolittle decomposition (0.0s in the strictly lower part)
            int rows = matrix.Length; int cols = matrix[0].Length;
            double[][] result = MatrixCreate(rows, cols);
            for (int i = 0; i < rows; ++i)
            {
                for (int j = 0; j < cols; ++j)
                {
                    if (i <= j)
                        result[i][j] = matrix[i][j];
                }
            }
            return result;
        }

        // --------------------------------------------------------------------------------------------------------------

        static double[][] PermArrayToMatrix(int[] perm)
        {
            // convert Doolittle perm array to corresponding perm matrix
            int n = perm.Length;
            double[][] result = MatrixCreate(n, n);
            for (int i = 0; i < n; ++i)
                result[i][perm[i]] = 1.0;
            return result;
        }

        static double[][] UnPermute(double[][] luProduct, int[] perm)
        {
            // unpermute product of Doolittle lower * upper matrix according to perm[]
            // no real use except to demo LU decomposition, or for consistency testing
            double[][] result = MatrixDuplicate(luProduct);

            int[] unperm = new int[perm.Length];
            for (int i = 0; i < perm.Length; ++i)
                unperm[perm[i]] = i;

            for (int r = 0; r < luProduct.Length; ++r)
                result[r] = luProduct[unperm[r]];

            return result;
        } // UnPermute

        // --------------------------------------------------------------------------------------------------------------
        //string[] GifaValues = ;

        static string VectorAsString(double[] vector)
        {

            string s = "";
            for (int i = 0; i < vector.Length; ++i)
            {
                s += vector[i].ToString("F3").PadLeft(8) + Environment.NewLine;
                //MessageBox.Show(vector[i].ToString("F3"));
            }

            s += Environment.NewLine;
            //MessageBox.Show(s);
            return s;
        }


        static string Gifa0_Values = "";
        static string Gifa1_Values = "";
        static string Gifa2_Values = "";
        static string Gifa3_Values = "";
        static string Gifa4_Values = "";
        static string Gifa5_Values = "";
        static string Gifa6_Values = "";
        static string Gifa7_Values = "";
        static string Gifa8_Values = "";
        static string GifaSum_Values = "";


        static double TotalOfGifaValues = 0.0;
        static double[] GifaValues = null;

        static string VectorAsStringGIFA(double[] vector)
        {
            string s = "";
            for (int i = 0; i < vector.Length; ++i)
            {
                s += vector[i].ToString("F3").PadLeft(8) + Environment.NewLine;
                
                //MessageBox.Show(vector[i].ToString("F3"));

            }
            s += Environment.NewLine;

            //MessageBox.Show( "Length: - " + vector.Length.ToString("F3"));

            GifaValues = new double[(int)(vector.Length - 5)];
            //double[] GifaValues = new double[(int)(vector.Length)];

            for (int i = 0; i < vector.Length - 5; ++i)
            {
                GifaValues[i] = vector[i]/1000;
                //EEGifaValues [i] = vector[i];

               //MessageBox.Show(GifaValues[i].ToString("F3"));
            }


            Gifa0_Values = GifaValues[0].ToString("n");
            Gifa1_Values = GifaValues[1].ToString("n");
            Gifa2_Values = GifaValues[2].ToString("n");
            Gifa3_Values = GifaValues[3].ToString("n");
            Gifa4_Values = GifaValues[4].ToString("n");
            Gifa5_Values = GifaValues[5].ToString("n");
            Gifa6_Values = GifaValues[6].ToString("n");
            Gifa7_Values = GifaValues[7].ToString("n");
            Gifa8_Values = GifaValues[8].ToString("n");
            
            //MessageBox.Show(TotalOfGifaValues.ToString("F3")); 

            TotalOfGifaValues = 0.0;
            //EEGifaValues = GifaValues;

            for (int i = 0; i < vector.Length - 5; ++i)
                TotalOfGifaValues += GifaValues[i];

            // MessageBox.Show(TotalOfGifaValues.ToString("F3")); 
            GifaSum_Values = TotalOfGifaValues.ToString("n");


            return s;
        }


        //static void procedure(double [] GifaList )
        //    {
        //    int i;
        //    for (i=0; i< GifaList.Length; i++)
        //    myListBox.Items.Add(aList[i].ToString()); // This will fill ListBox with
        //    strings of the double values contained in the ArrayList aList

        //    }


        static string VectorAsString(int[] vector)
        {
            string s = "";
            for (int i = 0; i < vector.Length; ++i)
                s += vector[i].ToString().PadLeft(2) + " ";
            s += Environment.NewLine;
            //MessageBox.Show(s);
            return s;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //var menuGroup = (from DataRow row1 in _ds.Tables["Rest_grdvitems"].Rows
            //                 orderby row1["Menu_Group"]
            //                 select row1["Menu_Group"].ToString()).Distinct();


            //var menuGroup = (from DataRow row1 in dtM[0].Tables["Rest_grdvitems"].Rows
            //                 orderby row1["Menu_Group"]
            //                 select row1["Menu_Group"].ToString()).Distinct();


            //TreeNode node = new TreeNode("All Items");


            //treeView1.BeginUpdate();

            //treeView1.Nodes.Add(node);
            //foreach (string menuitem in menuGroup)
            //{
            //    TreeNode node1 = new TreeNode(menuitem);
            //    TV_Categories_List.Nodes.Add(node1);
            //}

            //treeView1.EndUpdate();

            //TV_Categories_List.BeginUpdate();

            //TV_Categories_List.Nodes.Add(node);
            //foreach (string menuitem in menuGroup)
            //{
            //    TreeNode node1 = new TreeNode(menuitem);
            //    TV_Categories_List.Nodes.Add(node1);
            //}

            //TV_Categories_List.EndUpdate();
        }



        private void treeView2_AfterSelect(object sender, TreeViewEventArgs e)
        {

            // Retrieve data from the current TreeNode
            TreeNode node = e.Node;
            //DataGridViewRow row = (DataGridViewRow)node.Tag;

            //row.Selected = true;
            GifaDataGridView.ClearSelection();

            String searchValue = node.Text;
            int rowIndex = -1;

            foreach (DataGridViewRow Grow in GifaDataGridView.Rows)
            {
                if (Grow.Cells[14].Value.ToString().Equals(searchValue))
                {
                    rowIndex = Grow.Index;
                    Grow.Selected = true;

                    if (rowIndex > 0)
                        GifaDataGridView.FirstDisplayedScrollingRowIndex = GifaDataGridView.SelectedRows[0].Index - 1;
                    else
                        GifaDataGridView.FirstDisplayedScrollingRowIndex = GifaDataGridView.SelectedRows[0].Index;


                    break;
                }
            }

            ////currentIndex is the index of grid row
            //var rowElement = this.getView().getRecord(currentIndex);
            //this.getView().focusRow(rowElement);


            // GifaDataGridView.ClearSelection();
        }

        private void treeView2_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {

            //for (int i = 0; i < treeView2.Nodes.Count; i++)
            //{
            //    TreeNode node = treeView2.Nodes[i];
            //    DataGridViewRow row = GifaDataGridView.Rows[i];
            //    node.Tag = row;
            //}
            //for (int i = 0; i < treeView2.Nodes.Count; i++)
            //{
            //    TreeNode node = treeView2.Nodes[i];
            //    DataGridViewRow row = GifaDataGridView.Rows[i];
            //    node.Tag = row;
            //}
            //GifaDataGridView.ClearSelection();

        }

        DataTable treeViewItemsTable = null;
        //DataRow ItemsRow;
        private void PrintRecursive(TreeNode treeNode)
        {

            DataRow ItemsRow = treeViewItemsTable.NewRow();

            ItemsRow[0] = treeNode.Text;

            treeViewItemsTable.Rows.Add(ItemsRow);



            foreach (TreeNode tn in treeNode.Nodes)
            {
                PrintRecursive(tn);

            }

            // dataGridView1.DataSource = treeViewItemsTable;
        }

        // Call the procedure using the TreeView.
        private void CallRecursive(TreeView treeView)
        {
            // Print each node recursively.
            TreeNodeCollection nodes = treeView.Nodes;
            foreach (TreeNode n in nodes)
            {
                PrintRecursive(n);

                //TreeNode node = treeView2.Nodes[i];
                foreach (DataGridViewRow row in GifaDataGridView.Rows)
                {
                    if (row.IsNewRow) continue;
                    //row = GifaDataGridView.Rows[i];
                    n.Tag = row;
                }

            }



        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            // treeViewItemsTable = new DataTable();
            // treeViewItemsTable.Columns.Add("Number", typeof(string));
            // treeViewItemsTable.Columns.Add("Item", typeof(string));


            // CallRecursive( treeView2 );

            // string ExcelInputFolder2 = "C:\\Users\\p0077247\\documents\\Visual Studio 2010\\Projects\\Embodied Carbon Analysis\\NRMTemplate2.xlsx";

            // string ExcelfilePath2 = ExcelInputFolder2;//

            // getExcelData(ExcelfilePath2);
            // converToCSV3(0);

            // int NodeCount = treeView2.Nodes.Count;
            // int NodeCount2 = treeView2.GetNodeCount(true);

            //MessageBox.Show(NodeCount2.ToString());


            //for (int i = 0; i < treeView2.Nodes.Count; i++)
            //{
            //    TreeNode node = treeView2.Nodes[i];
            //    DataGridViewRow row = GifaDataGridView.Rows[i];
            //    node.Tag = row;
            //}








        }




        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

            BackgroundWorker worker = sender as BackgroundWorker;

            for (int i = 1; i <= 10; i++)
            {
                if (worker.CancellationPending == true)
                {
                    e.Cancel = true;
                    break;
                }
                else
                {

                    worker.ReportProgress(i);
                    //backgroundWorker1.WorkerReportsProgress = true;

                    //backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);


                    // Perform a time consuming operation and report progress.
                    RptGetDatasetElem();
                    //System.Threading.Thread.Sleep(500);
                    //worker.ReportProgress(i * 10);


                    // Simulate long task
                    System.Threading.Thread.Sleep(100);
                }
            }
        }

        

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                //resultLabel.Text = "Canceled!";
            }
            else if (e.Error != null)
            {
                // resultLabel.Text = "Error: " + e.Error.Message;
            }
            else
            {
                //resultLabel.Text = "Done!";
            }
        }

        private void iCEItemTableBindingNavigatorSaveItem_Click_1(object sender, EventArgs e)
        {
            this.Validate();
            this.iCEItemTableBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.iCEDatabaseDataSet);

        }

        private void Embodied_Energy_and_Carbon_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'iCEDatabaseDataSet.ICEItemTable' table. You can move, or remove it, as needed.
            this.iCEItemTableTableAdapter.Fill(this.iCEDatabaseDataSet.ICEItemTable);
            // TODO: This line of code loads data into the 'iCEDatabaseDataSet.ICEItemTable' table. You can move, or remove it, as needed.
            this.iCEItemTableTableAdapter.Fill(this.iCEDatabaseDataSet.ICEItemTable);
            // TODO: This line of code loads data into the 'iCEDatabaseDataSet.ICEItemTable' table. You can move, or remove it, as needed.
            this.iCEItemTableTableAdapter.Fill(this.iCEDatabaseDataSet.ICEItemTable);

            //dataGridView1.DataSource = M_QuantitiesTableValue;

        }

        private void iCEItemTableBindingNavigatorSaveItem_Click_2(object sender, EventArgs e)
        {
            this.Validate();
            this.iCEItemTableBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.iCEDatabaseDataSet);

        }

        private void iCEItemTableBindingNavigatorSaveItem_Click_3(object sender, EventArgs e)
        {
            this.Validate();
            this.iCEItemTableBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.iCEDatabaseDataSet);

        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            TransferQuantitiesToGifaGridView();
        }

        private void ExitButton_Click(object sender, EventArgs e)
        {
            EmbodiedEnergyCarbonFormExit();
        }

        public void EmbodiedEnergyCarbonFormExit()
        {


            //m_dataBuffer.ElementInformation.Dispose();
            //m_dataBuffer.ColElementInformation.Dispose();

            this.Close();


        }

        private System.Windows.Forms.DataVisualization.Charting.Chart chart4;
        string checkedProject = "";


        //private void CIOptionsDesignRadioButton_CheckedChanged(object sender, EventArgs e)
        //{

        //    string chartTitle = "";

        //    if (sender == CIOptionsDesignRadioButton)
        //    {

        //        chartTitle = CIOptionsDesignRadioButton.Text + ", " + ProjectRecordPeriod;

        //        // GetProjectPerformanceData();
        //        CIPlotGenericSingleChart();
        //        CIPlotGenericSingleChart0();
        //        //chart2.Titles.Add(chartTitle); 

        //        // Add Chart Titles
        //        chart4.Titles.Add(chartTitle);
        //        //chart1.Titles.Add("Title_2");
        //        //chart1.Titles.Add("Title_3");                

        //        // Set Title FontStyle
        //        chart4.Titles[0].Font = new Font("Microsoft Sans Serif", 11, FontStyle.Bold);

        //    }

        //}

        private void PlotGenericCharts()
        {

            string chartTitle = "Building Embodied Carbon Summary";

            //if (sender == CIOptionsDesignRadioButton)
            //{

                //chartTitle = CIOptionsDesignRadioButton.Text + ", " + ProjectRecordPeriod;

                // GetProjectPerformanceData();
                CIPlotGenericSingleChart();
                CIPlotGenericSingleChart0();
                //chart2.Titles.Add(chartTitle); 

                // Add Chart Titles
                chart4.Titles.Add(chartTitle);
                //chart1.Titles.Add("Title_2");
                //chart1.Titles.Add("Title_3");                

                // Set Title FontStyle
                chart4.Titles[0].Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);

            //}

        }

        private void CIPlotGenericSingleChart()
        {
            
            checkedProject = " Embodied Carbon ";
            //string chartTitle = "";

            //chartTitle = HeatingRadioButton.Text;

            //GetProjectPerformanceData();

            GifaSumaryGroupBox.Controls.Remove(chart4);

            // Create a Chart
            chart4 = new Chart();

            // Create Chart Area
            ChartArea chartArea1 = new ChartArea();

            // Add Chart Area to the Chart
            chart4.ChartAreas.Add(chartArea1);

            ////// Create a data series            
            //Series series1 = new Series();

            ////// Create a data series            
            Series series1 = new Series(checkedProject);

            chart4.Series.Add(series1);


            // PlotGenericSingleChart1();

            // Initializes a new instance of the DataSet class
            //DataSet myDataSet1 = new DataSet();


            ////myDataSet1.Tables.Add(dtM[3]);

            //string[] ColHeadings = new string[dtM[3].Columns.Count];

            //string[] RowValues = new string[dtM[3].Columns.Count];

            //DataRow row1 = dtM[3].Rows[0];
            ////DataRow row2 = dtM[3].Rows[1];

            //// DataRow row = dtM[3].Rows[1];
            //for (int colIndex = 1; colIndex < dtM[3].Columns.Count - 1; colIndex++)
            //{

            //    //ColHeadings[colIndex] = row1[colIndex].ToString();

            //    ColHeadings[colIndex] = dtM[3].Columns[colIndex].ColumnName;

            //    RowValues[colIndex] = row1[colIndex].ToString();


            //}



            //// now iterate through the arrays to add points to the "ByPoint" series,
            ////  setting X and Y values
            //for (int i = 1; i < ColHeadings.Length-1; i++)
            //{

            //    double YVal = Convert.ToDouble(RowValues[i]);

            //    //MessageBox.Show(ColHeadings[i] + " oti " + YVal.ToString());

            //    //chart1.ChartAreas["ChartArea1"].AxisX.MinorGrid.Enabled = true;
            //    chart2.Series["Series1"].Points.AddXY(ColHeadings[i], YVal);
            //    //chart1.Series["Series1"].Points.AddXY(FreqCostSet[i], ProbSet[i]);
            //}



            ////chart2.Series.Add(seriesName);
            //chart2.Series["Series1"].ChartType = SeriesChartType.Line;

            chart4.Series[checkedProject].BorderWidth = 2;



            // Show data points values as labels
            chart4.Series[checkedProject].IsValueShownAsLabel = false;

            // Set series point labels format
            chart4.Series[checkedProject].LabelFormat = "F0";

            chart4.Series[checkedProject].LabelBackColor = Color.Beige;

            // Set Border Color
            chart4.Series[checkedProject].LabelBorderColor = Color.Blue;

            // Set Border Style
            chart4.Series[checkedProject].LabelBorderDashStyle = ChartDashStyle.Solid;

            // Set Border Width
            chart4.Series[checkedProject].LabelBorderWidth = 1;

            chart4.Series[checkedProject].Font = new Font("Microsoft Sans Serif", 8, FontStyle.Bold);

            // Set marker attributes for the whole series            
            chart4.Series[checkedProject].MarkerStyle = MarkerStyle.Circle;
            chart4.Series[checkedProject].MarkerSize = 4;
            chart4.Series[checkedProject].MarkerColor = Color.Magenta;
            chart4.Series[checkedProject].MarkerBorderColor = Color.Red;
            chart4.Series[checkedProject].MarkerBorderWidth = 1;

            // Disable legend item for the first series
            chart4.Series[checkedProject].IsVisibleInLegend = true;





            //// Add a second legend
            //Legend legend1 = new Legend();

            ////// Set legend docking
            ////chart2.Legends["legend1"].Docking = Docking.Right;

            ////// Set legend alignment
            ////chart2.Legends["legend1"].Alignment = StringAlignment.Center;

            ////this.Chart1.Legends.Add(secondLegend);
            //this.chart2.Legends.Add(legend1);


            //// Set legend docking
            //chart2.Legends["legend1"].Docking = Docking.Right;

            //// Set legend alignment
            //chart2.Legends["legend1"].Alignment = StringAlignment.Center;

            //// Set legend position
            //chart2.Legend.Position.Auto = false;
            //chart2.Legend.Position.X = 35;
            //chart2.Legend.Position.Y = 40;
            //chart2.Legend.Position.Width = 35;
            //chart2.Legend.Position.Height = 10;
            string Ytitle = ""; 

            if(EERadioButton.Checked)
              Ytitle = "Embodied Energy (GJ)";
            
            if(ECRadioButton.Checked)
               Ytitle = "Embodied Carbon (tCO2)";

            //if (CISingleProjectCheckBox.Checked)
            //{
            //    AnyGrossArea = 1;
            //    Ytitle = "Energy Cost (£)";
            //}

            //if (CICombineProjectsCheckBox.Checked)
            //{
            //    AnyGrossArea = BuildingGrossArea;
            //    Ytitle = "Energy Cost/gross bldg area (£/GBA)";
            //}



            chart4.BorderSkin.SkinStyle = BorderSkinStyle.Raised;

            ///////////////////////////////////////////////


            //////Background area appearance

            // Set Back Color
            chart4.BackColor = Color.LightGreen;

            // Set Back Gradient End Color
            chart4.BackSecondaryColor = Color.LightSkyBlue;

            // Set Hatch Style
            chart4.BackHatchStyle = ChartHatchStyle.DashedHorizontal;

            // Set Gradient Type
            chart4.BackGradientStyle = GradientStyle.DiagonalRight;

            // Set Border Color
            chart4.BorderColor = Color.Blue;

            // Set Border Style
            chart4.BorderDashStyle = ChartDashStyle.Solid;

            // Set Border Width
            chart4.BorderWidth = 1;

            // Chart Image Mode
            chart4.BackImageWrapMode = ChartImageWrapMode.TileFlipX;

            //Chart Image Align
            //chart1.BackImageAlignment = ChartImageAlignStyle.BottomLeft;
            chart4.BackImageAlignment = ChartImageAlignmentStyle.BottomLeft;

            //// Set Image
            //chart1.BackImage = "Brain.jpg";        


            chart4.BorderSkin.BackColor = Color.CadetBlue;

            // Set Back Gradient End Color
            chart4.BorderSkin.BackSecondaryColor = Color.Blue;

            // Set Hatch Style
            chart4.BorderSkin.BackHatchStyle = ChartHatchStyle.DarkVertical;

            // Set Gradient Type
            chart4.BorderSkin.BackGradientStyle = GradientStyle.DiagonalRight;

            // Set Border Color
            chart4.BorderSkin.BorderColor = Color.Yellow;

            // Set Border Style
            chart4.BorderSkin.BorderDashStyle = ChartDashStyle.Solid;

            // Set Border Width
            chart4.BorderSkin.BorderWidth = 2;


            //////Axis line appearance
            // Set Axis Color
            chart4.ChartAreas["ChartArea1"].AxisY.LineColor = Color.Blue;

            // Set Axis Line Style
            chart4.ChartAreas["ChartArea1"].AxisY.LineDashStyle = ChartDashStyle.Solid;

            // Set Arrow Style
            chart4.ChartAreas["ChartArea1"].AxisY.ArrowStyle = AxisArrowStyle.None;

            // Set Line Width
            chart4.ChartAreas["ChartArea1"].AxisY.LineWidth = 1;

            chart4.ChartAreas["ChartArea1"].AxisX.LineColor = Color.Blue;

            // Set Axis Line Style
            chart4.ChartAreas["ChartArea1"].AxisX.LineDashStyle = ChartDashStyle.Solid;

            // Set Arrow Style
            chart4.ChartAreas["ChartArea1"].AxisX.ArrowStyle = AxisArrowStyle.None;

            // Set Line Width
            chart4.ChartAreas["ChartArea1"].AxisX.LineWidth = 1;

            // Set axis labels font
            chart4.ChartAreas["ChartArea1"].AxisX.LabelStyle.Font = new Font("Arial", 9);
            chart4.ChartAreas["ChartArea1"].AxisY.LabelStyle.Font = new Font("Arial", 9);

            // Set axis title
            chart4.ChartAreas["ChartArea1"].AxisX.Title = "Work breakdown Structure";
            //chart4.ChartAreas["ChartArea1"].AxisY.Title = "CO2 Emission (kgCO2)";
            chart4.ChartAreas["ChartArea1"].AxisY.Title = Ytitle;


            // Set Title font
            chart4.ChartAreas["ChartArea1"].AxisX.TitleFont = new Font("Microsoft Sans Serif", 9, FontStyle.Regular);
            chart4.ChartAreas["ChartArea1"].AxisY.TitleFont = new Font("Microsoft Sans Serif", 9, FontStyle.Regular);

            // Set Title color
            chart4.ChartAreas["ChartArea1"].AxisX.TitleForeColor = Color.Black;
            chart4.ChartAreas["ChartArea1"].AxisX.TitleForeColor = Color.Black;

            // Enable X axis labels automatic fitting
            chart4.ChartAreas["ChartArea1"].AxisX.IsLabelAutoFit = true;
            chart4.ChartAreas["ChartArea1"].AxisY.IsLabelAutoFit = true;

            // Set X axis automatic fitting style
            chart4.ChartAreas["ChartArea1"].AxisX.LabelAutoFitStyle =
               LabelAutoFitStyles.DecreaseFont | LabelAutoFitStyles.IncreaseFont | LabelAutoFitStyles.WordWrap;
            chart4.ChartAreas["ChartArea1"].AxisY.LabelAutoFitStyle =
               LabelAutoFitStyles.DecreaseFont | LabelAutoFitStyles.IncreaseFont | LabelAutoFitStyles.WordWrap;


            //// Enable X axis labels automatic fitting
            //Chart1.ChartAreas["Default"].AxisX.IsLabelAutoFit = true;

            //// Set X axis automatic fitting style
            //Chart1.ChartAreas["Default"].AxisX.LabelAutoFitStyle =
            //    LabelAutoFitStyle.DecreaseFont | LabelAutoFitStyle.IncreaseFont | LabelAutoFitStyle.WordWrap;


            //// Add Chart Titles
            //chart2.Titles.Add(chartTitle);
            ////chart1.Titles.Add("Title_2");
            ////chart1.Titles.Add("Title_3");




            //// Set Title FontStyle
            //chart2.Titles[0].Font = new Font("Microsoft Sans Serif", 11, FontStyle.Bold);
            ////chart1.Titles[0].BackColor = Color.BlueViolet;



            //legend1 = new Legend();
            ////this.Chart1.Legends.Add(secondLegend);
            //this.chart4.Legends.Add(legend1);
            ////this.chart2.Legends["Legend1"].Enabled = true;



            chart4.BorderSkin.SkinStyle = BorderSkinStyle.Raised;


            // Set chart control location       
            chart4.Location = new System.Drawing.Point(478, 50);


            // Set Chart control size
            //chart1.Size = new System.Drawing.Size(690, 440);
            chart4.Size = new System.Drawing.Size(380, 276);

            /////////////////////////////////////////////////
            //// Set legend docking
            //chart2.Legends["legend1"].Docking = Docking.Right;

            //// Set legend alignment
            //chart2.Legends["legend1"].Alignment = StringAlignment.Center;


            // PerformanceGroupBox.Controls.AddRange(new System.Windows.Forms.Control[] { chart1, chart2 });
            GifaSumaryGroupBox.Controls.AddRange(new System.Windows.Forms.Control[] { chart4 });
        }


        private void CIPlotGenericSingleChart0()
        {


            

            string EE_0 = GifaDataGridView[10, Gifa0_RowToColIndex].Value.ToString();
            string EE_1 = GifaDataGridView[10, Gifa1_RowToColIndex].Value.ToString();
            string EE_2 =GifaDataGridView[10, Gifa2_RowToColIndex].Value.ToString();
            string EE_3 =GifaDataGridView[10, Gifa3_RowToColIndex].Value.ToString();
            string EE_4 =GifaDataGridView[10, Gifa4_RowToColIndex].Value.ToString();
            string EE_5 =GifaDataGridView[10, Gifa5_RowToColIndex].Value.ToString();
            string EE_6 =GifaDataGridView[10, Gifa6_RowToColIndex].Value.ToString();
            string EE_7 =GifaDataGridView[10, Gifa7_RowToColIndex].Value.ToString();
            string EE_8 = GifaDataGridView[10, Gifa8_RowToColIndex].Value.ToString();

            string EC_0 = GifaDataGridView[11, Gifa0_RowToColIndex].Value.ToString();
            string EC_1 = GifaDataGridView[11, Gifa1_RowToColIndex].Value.ToString();
            string EC_2 = GifaDataGridView[11, Gifa2_RowToColIndex].Value.ToString();
            string EC_3 = GifaDataGridView[11, Gifa3_RowToColIndex].Value.ToString();
            string EC_4 = GifaDataGridView[11, Gifa4_RowToColIndex].Value.ToString();
            string EC_5 = GifaDataGridView[11, Gifa5_RowToColIndex].Value.ToString();
            string EC_6 = GifaDataGridView[11, Gifa6_RowToColIndex].Value.ToString();
            string EC_7 = GifaDataGridView[11, Gifa7_RowToColIndex].Value.ToString();
            string EC_8 = GifaDataGridView[11, Gifa8_RowToColIndex].Value.ToString();


            string[] GifaHeadings = new string[] {
            GifaDataGridView[0, Gifa0_RowToColIndex].Value.ToString(),
            GifaDataGridView[0, Gifa1_RowToColIndex].Value.ToString(),
            GifaDataGridView[0, Gifa2_RowToColIndex].Value.ToString(),
            GifaDataGridView[0, Gifa3_RowToColIndex].Value.ToString(),
            GifaDataGridView[0, Gifa4_RowToColIndex].Value.ToString(),
            GifaDataGridView[0, Gifa5_RowToColIndex].Value.ToString(),
            GifaDataGridView[0, Gifa6_RowToColIndex].Value.ToString(),
            GifaDataGridView[0, Gifa7_RowToColIndex].Value.ToString(),
            GifaDataGridView[0, Gifa8_RowToColIndex].Value.ToString()};

            double[] Gifa_EEValues = new double[] {
            Convert.ToDouble(EE_0),
            Convert.ToDouble(EE_1),
            Convert.ToDouble(EE_2),
            Convert.ToDouble(EE_3),
            Convert.ToDouble(EE_4),
            Convert.ToDouble(EE_5),
            Convert.ToDouble(EE_6),
            Convert.ToDouble(EE_7),
            Convert.ToDouble(EE_8)};


            double[] Gifa_ECValues = new double[] {
            Convert.ToDouble(EC_0),
            Convert.ToDouble(EC_1),
            Convert.ToDouble(EC_2),
            Convert.ToDouble(EC_3),
            Convert.ToDouble(EC_4),
            Convert.ToDouble(EC_5),
            Convert.ToDouble(EC_6),
            Convert.ToDouble(EC_7),
            Convert.ToDouble(EC_8)};

            if (EERadioButton.Checked)
            {

                // now iterate through the arrays to add points to the "ByPoint" series,
                //  setting X and Y values
                for (int i = 0; i < GifaHeadings.Length; i++)
                {

                    double YVal = Gifa_EEValues[i];

                    //MessageBox.Show(ColHeadings[i] + " oti " + YVal.ToString());

                    //chart1.ChartAreas["ChartArea1"].AxisX.MinorGrid.Enabled = true;
                    chart4.Series[checkedProject].Points.AddXY(GifaHeadings[i], YVal);
                    //chart1.Series["Series1"].Points.AddXY(FreqCostSet[i], ProbSet[i]);
                }
            }


            if (ECRadioButton.Checked)
            {
                // now iterate through the arrays to add points to the "ByPoint" series,
                //  setting X and Y values
                for (int i = 0; i < GifaHeadings.Length; i++)
                {

                    double YVal = Gifa_ECValues[i];

                    //MessageBox.Show(ColHeadings[i] + " oti " + YVal.ToString());

                    //chart1.ChartAreas["ChartArea1"].AxisX.MinorGrid.Enabled = true;
                    chart4.Series[checkedProject].Points.AddXY(GifaHeadings[i], YVal);
                    //chart1.Series["Series1"].Points.AddXY(FreqCostSet[i], ProbSet[i]);
                }
            }



            //chart2.Series.Add(seriesName);
            //chart3.Series[checkedProject].ChartType = SeriesChartType.StackedBar;
            //chart3.Series[checkedProject].ChartType = SeriesChartType.StackedBar;
            //chart3.Series[checkedProject].ChartType = SeriesChartType.Line;
            //chart3.Series[checkedProject].ChartType = SeriesChartType.StackedBar;
            //chart3.Series[checkedProject].ChartType = SeriesChartType.StackedColumn;
            chart4.Series[checkedProject].ChartType = SeriesChartType.Column;



        }
        
        private void EERadioButton_CheckedChanged(object sender, EventArgs e)
        {

            
            string chartTitle = "";

            if (sender == EERadioButton)
            {

                chartTitle = "Building Embodied Energy Summary";

                //chartTitle = CIOptionsDesignRadioButton.Text + ", " + ProjectRecordPeriod;

                // GetProjectPerformanceData();
                CIPlotGenericSingleChart();
                CIPlotGenericSingleChart0();
                //chart2.Titles.Add(chartTitle); 

                // Add Chart Titles
                chart4.Titles.Add(chartTitle);
                //chart1.Titles.Add("Title_2");
                //chart1.Titles.Add("Title_3");                

                // Set Title FontStyle
                chart4.Titles[0].Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);

            }
        }

        private void ECRadioButton_CheckedChanged(object sender, EventArgs e)
        {


            string chartTitle = "";

            if (sender == ECRadioButton)
            {

                chartTitle = "Building Embodied Carbon Summary";

                //chartTitle = CIOptionsDesignRadioButton.Text + ", " + ProjectRecordPeriod;

                // GetProjectPerformanceData();
                CIPlotGenericSingleChart();
                CIPlotGenericSingleChart0();
                //chart2.Titles.Add(chartTitle); 

                // Add Chart Titles
                chart4.Titles.Add(chartTitle);
                //chart1.Titles.Add("Title_2");
                //chart1.Titles.Add("Title_3");                

                // Set Title FontStyle
                chart4.Titles[0].Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);

            }

        }



        

    }
}
