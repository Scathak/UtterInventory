using Excel = Microsoft.Office.Interop.Excel;

namespace UtterInventory
{
    public partial class ThisAddIn
    {
        public bool OnStartEmpty; 
        public object[][] strings = new object[2][];
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            OnStartEmpty = IsEmpty();
            this.Application.WorkbookBeforeSave += new Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
            this.Application.WorkbookAfterSave += new Excel.AppEvents_WorkbookAfterSaveEventHandler(Application_WorkbookAfterSave);
        }
        void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool Success1, ref bool Success)
        {
            this.Application.DisplayAlerts = false;
            foreach (Excel.Worksheet sh in Wb.Sheets)
            {
                if(sh.Name != "_rawData")
                {
                    sh.Delete();
                }
            }
            this.Application.DisplayAlerts = true;
        }
        void Application_WorkbookAfterSave(Excel.Workbook Wb, bool Success)
        {
            DeployTables(7, 1);
        }
        bool IsEmpty() {
            return (0 == this.Application.Workbooks.Count);
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
