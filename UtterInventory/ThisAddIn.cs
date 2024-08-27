using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using System.Threading;

namespace UtterInventory
{
    public partial class ThisAddIn
    {
        public bool bPlayflag = false;
        public Thread ThreadCam;
        public string ipCameraAddress = string.Empty;
        private CustomTaskPane myCustomTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave += new AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
            this.Application.WorkbookAfterSave += new AppEvents_WorkbookAfterSaveEventHandler(Application_WorkbookAfterSave);
            this.Application.WorkbookOpen += AppEvents_WorkbookOpen;

            UserContropPanel_Init();
        }
        private void UserContropPanel_Init()
        {
            // Create an instance of the User Control
            var myUserControl = new UserControl1();

            // Add the User Control to a Custom Task Pane
            myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl, "Camera panel");

            // Set the properties of the Custom Task Pane
            myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            myCustomTaskPane.Width = 300;
            myCustomTaskPane.Visible = false;
        }
        private void AppEvents_WorkbookOpen(Workbook Wb)
        {
            Globals.ThisAddIn.totalNumberOfStyles = Wb.TableStyles.Count;
            if (workSheetExist(Wb, rawDataSheetName))
            {
                RefreshCache(Wb.Sheets[rawDataSheetName]);
            }
            Globals.ThisAddIn.Application.CutCopyMode = XlCutCopyMode.xlCut;
        }
        // Method to toggle the visibility of the task pane
        public void ToggleTaskPaneVisibility(bool isVisible)
        {
            if (myCustomTaskPane != null)
            {
                myCustomTaskPane.Visible = isVisible;
            }
        }
        void Application_WorkbookBeforeSave(Workbook Wb, bool Success1, ref bool Success)
        {
            this.Application.DisplayAlerts = false;
            this.Application.ScreenUpdating = false;
            Wb.Sheets[rawDataSheetName].Visible = XlSheetVisibility.xlSheetVisible;

            foreach (Worksheet sh in Wb.Sheets)
            {
                if(sh.Name != rawDataSheetName)
                {
                    sh.Delete();
                }
            }
            this.Application.DisplayAlerts = true;
            this.Application.ScreenUpdating = true;
        }
        void Application_WorkbookAfterSave(Workbook Wb, bool Success)
        {
            this.Application.DisplayAlerts = false;
            this.Application.ScreenUpdating = false;
            DeployTables(Wb, topLeftCornerTableRow, topLeftCornerTableCol);
            RefreshCache(Wb.Sheets[rawDataSheetName]);
            Globals.ThisAddIn.DataReplication(Wb, topLeftCornerTableRow, topLeftCornerTableCol);
            this.Application.DisplayAlerts = true;
            this.Application.ScreenUpdating = true;
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
