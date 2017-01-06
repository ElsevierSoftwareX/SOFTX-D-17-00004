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
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
//using AnalyticalSupportData_info;

using System.Data.OleDb;
using System.Xml;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using AnalyticalSupportData_Info;
using AnalyticalSupportData_info;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using System.Drawing.Drawing2D;
using System.Threading;
using System.Diagnostics;
//#include <windows.h>
//#include <stdio.h>
//#include <stdlib.h>


using System.Globalization;
using System.Text.RegularExpressions;
using System.Collections;
using System.Configuration;




namespace Revit.SDK.Samples.AnalyticalSupportData_Info.CS
{
    /// <summary>
    /// UI which display the information
    /// </summary>
    //public partial class OperationMode : System.Windows.Forms.Form
    public partial class OperationMode : System.Windows.Forms.Form
    {
        // an instance of Command class which is prepared the displayed data.
        Command m_dataBuffer;
       // TextBox projectIDTextBox;
        /// <summary>
        /// Default constructor
        /// </summary>
        OperationMode()
        {
            InitializeComponent();
            //m_elementInformation = StoreInformationInDataTable(Pline);
        }
        //public OperationMode(TextBox ProjectIDTextBox)
        //    : this()
        //{
        //    projectIDTextBox = ProjectIDTextBox;
        //}
        /// <summary>
        /// constructor
        /// </summary>
        /// <param name="dataBuffer"></param>
        public OperationMode(Command dataBuffer) : this()
        {
            m_dataBuffer = dataBuffer;
            ProceedButton.Enabled = false;
           
        }


             
        


        //public virtual void Unlock(long position, long length); 
       
        /// <summary>
        /// exit
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void closeButton_Click(object sender, EventArgs e)
        {
            this.Close();
            //m_elementInformation = StoreInformationInDataTable(Pline);           
           
                       
        }

        
       


        public RadioButton EECAnalyisisRadioButton
        {

            get { return EmbodiedECAnalyisisRadioButton; }
        }



        public string ProjIDTextBox
        {

            get { return ProjectIDTextBox.Text; }
        }


       

        public void ProceedButton_Click(object sender, EventArgs e)
        {

           // m_elementInformation = StoreInformationInDataTable(Pline);
           

            if (EmbodiedECAnalyisisRadioButton.Checked)
            {
               
                
               
                using (Embodied_Energy_and_Carbon f = new Embodied_Energy_and_Carbon() )
                {
                    

                    f.ShowDialog();

                }
            }

           
                           

        }



             




        private void EmbodiedECAnalyisisRadioButton_CheckedChanged(object sender, EventArgs e)
        {

            if (sender == EmbodiedECAnalyisisRadioButton)
                MessageBox.Show("Please ensure that 3D model is open and is the active window ", "Embodied Energy and Carbon Analysis",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);


            ProceedButton.Enabled = true;

        }




        /// <summary>
        /// ////////////////
        /// </summary>

       // String fileName;
       // String FileName;


/// <summ






   /// <summary>
   /// ////////////////////////////////////////////////
   /// </summary>
         


        protected int TextBoxCount = 5; // number of TextBoxes on Form

        // enumeration constants specify TextBox indices
        public enum TextBoxIndices
        {
            ProjectID,
            ProjectTitle,
            ProjectLocation,
            DesignOptionNo,
            DesignLife,
            
        } // end enum


        

        public void SetTextBoxValues(string[] values)
        {
            // determine whether string array has correct length
            if (values.Length != TextBoxCount)
            {
                // throw exception if not correct length
                throw (new ArgumentException("There must be " +
                   (TextBoxCount + 1) + " strings in the array"));
            } // end if
            // set array values if array has correct length
            else
            {
                // set array values to text box values

                ProjectIDTextBox.Text = values[(int)TextBoxIndices.ProjectID];
                ProjectTitleTextBox.Text = values[(int)TextBoxIndices.ProjectTitle];
                ProjectLocationTextBox.Text = values[(int)TextBoxIndices.ProjectLocation];

                DesignOptionNoTextBox.Text = values[(int)TextBoxIndices.DesignOptionNo];
                DesignLifeTextBox.Text = values[(int)TextBoxIndices.DesignLife];
                              


            } // end else
        } // end method SetTextBoxValues




        public string[] GetTextBoxValues()
        {
            string[] values = new string[TextBoxCount];

            // copy text box fields to string array
            values[(int)TextBoxIndices.ProjectID] = ProjectIDTextBox.Text;
            values[(int)TextBoxIndices.ProjectTitle] = ProjectTitleTextBox.Text;
            values[(int)TextBoxIndices.ProjectLocation] = ProjectLocationTextBox.Text;

            values[(int)TextBoxIndices.DesignOptionNo] = DesignOptionNoTextBox.Text;
            values[(int)TextBoxIndices.DesignLife] = DesignLifeTextBox.Text;


           
            return values;
        }

        private void OperationMode_Load(object sender, EventArgs e)
        {

           
        }

        
        
    }


}