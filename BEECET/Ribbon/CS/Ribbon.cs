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
using System.Linq;
using System.Text;
using Autodesk.Revit;
using System.Diagnostics;
using System.IO;
using System.Windows.Media;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI.Events;

namespace Revit.SDK.Samples.Ribbon.CS
{
    /// <summary>
    /// Implements the Revit add-in interface IExternalApplication,
    /// show user how to create RibbonItems by API in Revit.
    /// we add one RibbonPanel:
    /// 1. contains a SplitButton for user to create Non-Structural or Structural Wall;
    /// 2. contains a StackedButton which is consisted with one PushButton and two Comboboxes, 
    /// PushButton is used to reset all the RibbonItem, Comboboxes are use to select Level and WallShape
    /// 3. contains a RadioButtonGroup for user to select WallType.
    /// 4. Adds a Slide-Out Panel to existing panel with following functionalities:
    /// 5. a text box is added to set mark for new wall, mark is a instance parameter for wall, 
    /// Eg: if user set text as "wall", then Mark for each new wall will be "wall1", "wall2", "wall3"....
    /// 6. a StackedButton which consisted of a PushButton (delete all the walls) and a PulldownButton (move all the walls in X or Y direction)
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
   [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class Ribbon : IExternalApplication
    {
        // ExternalCommands assembly path
        static string AddInPath = typeof(Ribbon).Assembly.Location;
        // Button icons directory
        static string ButtonIconsFolder = Path.GetDirectoryName(AddInPath);
        // uiApplication
        //static UIApplication uiApplication = null;

        #region IExternalApplication Members
        /// <summary>
        /// Implement this method to implement the external application which should be called when 
        /// Revit starts before a file or default template is actually loaded.
        /// </summary>
        /// <param name="application">An object that is passed to the external application 
        /// which contains the controlled application.</param>
        /// <returns>Return the status of the external application. 
        /// A result of Succeeded means that the external application successfully started. 
        /// Cancelled can be used to signify that the user cancelled the external operation at 
        /// some point.
        /// If Failed is returned then Revit should inform the user that the external application 
        /// failed to load and the release the internal reference.</returns>
        public Autodesk.Revit.UI.Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // create customer Ribbon Items
                CreateRibbonSamplePanel(application);

                return Autodesk.Revit.UI.Result.Succeeded;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ribbon Sample");

                return Autodesk.Revit.UI.Result.Failed;
            }
        }

        /// <summary>
        /// Implement this method to implement the external application which should be called when 
        /// Revit is about to exit, Any documents must have been closed before this method is called.
        /// </summary>
        /// <param name="application">An object that is passed to the external application 
        /// which contains the controlled application.</param>
        /// <returns>Return the status of the external application. 
        /// A result of Succeeded means that the external application successfully shutdown. 
        /// Cancelled can be used to signify that the user cancelled the external operation at 
        /// some point.
        /// If Failed is returned then the Revit user should be warned of the failure of the external 
        /// application to shut down correctly.</returns>
        public Autodesk.Revit.UI.Result OnShutdown(UIControlledApplication application)
        {
            //remove events
            //List<RibbonPanel> myPanels = application.GetRibbonPanels();
            //Autodesk.Revit.UI.ComboBox comboboxLevel = (Autodesk.Revit.UI.ComboBox)(myPanels[0].GetItems()[2]);

            
           
            return Autodesk.Revit.UI.Result.Succeeded;
        }
        #endregion

        /// <summary>
        /// This method is used to create RibbonSample panel, and add wall related command buttons to it:
        /// 1. contains a SplitButton for user to create Non-Structural or Structural Wall;
        /// 2. contains a StackedBotton which is consisted with one PushButton and two Comboboxes, 
        /// PushButon is used to reset all the RibbonItem, Comboboxes are use to select Level and WallShape
        /// 3. contains a RadioButtonGroup for user to select WallType.
        /// 4. Adds a Slide-Out Panel to existing panel with following functionalities:
        /// 5. a text box is added to set mark for new wall, mark is a instance parameter for wall, 
        /// Eg: if user set text as "wall", then Mark for each new wall will be "wall1", "wall2", "wall3"....
        /// 6. a StackedButton which consisted of a PushButton (delete all the walls) and a PulldownButton (move all the walls in X or Y direction)
        /// </summary>
        /// <param name="application">An object that is passed to the external application 
        /// which contains the controlled application.</param>
        //string solutionDir;
        //string ProjectDirectory;
        private void CreateRibbonSamplePanel(UIControlledApplication application)
        {
            // create a Ribbon panel which contains three stackable buttons and one single push button.
            //string firstPanelName = "Ribbon Sample";
            string firstPanelName = "Sustainability Estimator";
            RibbonPanel ribbonSamplePanel = application.CreateRibbonPanel(firstPanelName);

           // ribbonSamplePanel.AddSeparator();

            #region Create a SplitButton for user to create Non-Structural or Structural Wall

            //PushButtonData pushButtonData = new PushButtonData("WallPush", "Sustainability", AddInPath, "Revit.SDK.Samples.Ribbon.CS.CreateWall");
            //PushButtonData deleteWallsButtonData = new PushButtonData("deleteWalls", "Delete Walls", AddInPath, "Revit.SDK.Samples.Ribbon.CS.DeleteWalls");

            //PushButtonData pushButtonData = new PushButtonData("SustEstimator", "Sustainability",
            //    @"C:\Users\ETONAKPO\Documents\Visual Studio 2010\Projects\AnaliticalSuppCS - Copy 9 - Copy (50) - Copy\bin\Debug\AnalyticalSupportData_Info.dll",
            //    "Revit.SDK.Samples.AnalyticalSupportData_Info.CS.Command");

            ///////////////////////////////////////////////////////
            // create and show dialog box enabling user to open file


           

            //string ProjectDirectory;

            //string ProjectDllDirectory = @"C:\Users\ETONAKPO\Documents\Visual Studio 2010\Projects\SteelSustainabilityEstimation\bin\Debug\AnalyticalSupportData_Info.dll";

           // MessageBox.Show(ProjectDllDirectory);

            PushButtonData pushButtonData = new PushButtonData("SustEstimator", "Sustainability",
                @" C:\Users\ETONAKPO\Documents\Visual Studio 2010\Projects\SteelSustainabilityEstimation\bin\Debug\AnalyticalSupportData_Info.dll",
                "Revit.SDK.Samples.AnalyticalSupportData_Info.CS.Command");

            //SplitButtonData splitButtonData = new SplitButtonData("SustEstimator", "Sustainability");

            //SplitButtonData splitButtonData = new SplitButtonData("SustEstimator", "Sustainability",
            //    @" C:\Users\ETONAKPO\Documents\Visual Studio 2010\Projects\SteelSustainabilityEstimation\bin\Debug\AnalyticalSupportData_Info.dll",
            //    "Revit.SDK.Samples.AnalyticalSupportData_Info.CS.Command");

            //PushButtonData pushButtonData = new PushButtonData("SustEstimator", "Sustainability",
            //   ProjectDllDirectory,
            //    "Revit.SDK.Samples.AnalyticalSupportData_Info.CS.Command");

           
            //PushButton pushButton = ribbonSamplePanel.AddItem(new PushButtonData("HelloWorld",
            //        "HelloWorld", @"D:\HelloWorld.dll", "HelloWorld.CsHelloWorld")) as PushButton;

            //PushButton pushButton = ribbonSamplePanel.AddItem(new PushButtonData("HelloWorld",
            //        "HelloWorld", @"C:\Users\ETONAKPO\Documents\Visual Studio 2010\Projects\AnaliticalSuppCS - Copy 9 - Copy (50) - Copy\bin\Debug\AnalyticalSupportData_Info.dll", "HelloWorld.CsHelloWorld")) as PushButton;

            //PushButton pushButton = ribbonSamplePanel.AddItem(pushButtonData) as PushButton;
            PushButton pushButton = ribbonSamplePanel.AddItem(pushButtonData) as PushButton;

            pushButton.LargeImage = new BitmapImage(new Uri(Path.Combine(ButtonIconsFolder, "CreateWall.png"), UriKind.Absolute));           
            pushButton.ToolTip = "Calls Steel Sustainability Estimator Programme.";
            pushButton.ToolTipImage = new BitmapImage(new Uri(Path.Combine(ButtonIconsFolder, "CreateWallTooltip.bmp"), UriKind.Absolute));
            
            #endregion

            
           
        }

       

    }
}
