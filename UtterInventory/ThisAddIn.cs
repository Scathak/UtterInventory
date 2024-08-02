using Microsoft.Office.Interop.Excel;

namespace UtterInventory
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave += new AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
            this.Application.WorkbookAfterSave += new AppEvents_WorkbookAfterSaveEventHandler(Application_WorkbookAfterSave);
            this.Application.WorkbookOpen += AppEvents_WorkbookOpen;
        }

        private void AppEvents_WorkbookOpen(Workbook Wb)
        {
            Globals.ThisAddIn.totalNumberOfStyles = Wb.TableStyles.Count;
            if (workSheetExist(Wb, rawDataSheetName))
            {
                RefreshCache(Wb.Sheets[rawDataSheetName]);
            }
        }

        void Application_WorkbookBeforeSave(Workbook Wb, bool Success1, ref bool Success)
        {
            this.Application.DisplayAlerts = false;
            foreach (Worksheet sh in Wb.Sheets)
            {
                if(sh.Name != rawDataSheetName)
                {
                    sh.Delete();
                }
            }
            this.Application.DisplayAlerts = true;
        }
        void Application_WorkbookAfterSave(Workbook Wb, bool Success)
        {
            DeployTables(Wb, topLeftCornerTableRow, topLeftCornerTableCol);
            RefreshCache(Wb.Sheets[rawDataSheetName]);
            Globals.ThisAddIn.DataReplication(Wb, topLeftCornerTableRow, topLeftCornerTableCol);
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
