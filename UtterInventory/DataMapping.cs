using Microsoft.Office.Interop.Excel;
using System.Linq;
using System;

namespace UtterInventory
{
    public partial class ThisAddIn
    {
        public void DataReplication(Workbook wb, int row, int col)
        {
            if (!workSheetExist(wb, rawDataSheetName)) { return; }

            Worksheet rawDataSheet = wb.Sheets[rawDataSheetName];
            rawDataSheet.Activate();
            Range lastCell = GetOccupiedCells(wb.ActiveSheet);
            Range OccupiedDataRange = rawDataSheet.Range[rawDataSheet.Cells[1, 1], lastCell];

            if (OccupiedDataRange.Rows.Count <= 1) return;

            var tableWidth = RawDataCache.GetLength(1);
            var columnHeigh = RawDataCache.GetLength(0);
            var tableHeaders = Enumerable.Range(1, tableWidth).Select(x => RawDataCache[1, x]).ToArray();

            foreach (var table in TablesStructure)
            { 
                Worksheet wsForCopy = null;
                if (GetTablesToCopyOn().ContainsKey(table.Key))
                {
                    wsForCopy = wb.Worksheets[GetAllTablesNames()[table.Key]];
                }else continue;
                var tableToCopyWidth = table.Value.Length;
                var colNames = table.Value.ToArray();
                var j = 0;
                object[,] cacheToCopy = new object[columnHeigh - 1, tableToCopyWidth];
                foreach (var column in colNames)
                {
                    var IndexOccurence = Array.FindIndex(tableHeaders, w => w.Equals(column));
                    if (IndexOccurence >= 0)
                    {
                        for (var rowId = 0; rowId < columnHeigh - 1; rowId++) {
                            cacheToCopy[rowId, j] = RawDataCache[rowId+2, IndexOccurence + 1];
                        }
                        j++;
                    }
                }
                var rangeToCopy = wsForCopy.ListObjects[wsForCopy.Name].Range;
                rangeToCopy.Offset[1, 0].Resize[columnHeigh-1].Cells.Value2 = cacheToCopy;
                SelectOneCell(wsForCopy, rangeToCopy,2,1);
            }
            rawDataSheet.Visible = XlSheetVisibility.xlSheetHidden;

            CreateQRcodesForTable(wb.Sheets[barCodeSheetName]);
        }
        public void ApplyColumnsTypes(Range OccupiedDataRange )
        {

        }
        public void SelectOneCell(Worksheet ws, Range range, int row, int col)
        {
            ws.Activate();
            range.Cells[row, col].Select();
        }
        public Range GetOccupiedCells(Worksheet ws)
        {
            ws.Activate();
            return ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
        }
        public static void CreateQRcodesForTable(Worksheet ws)
        {
            if (ws.Name.Equals(barCodeSheetName, StringComparison.OrdinalIgnoreCase))
            {
                Globals.ThisAddIn.Application.Application.DisplayAlerts = false;
                Globals.ThisAddIn.Application.Application.ScreenUpdating = false;

                Range usedRange = ws.ListObjects[ws.Name].Range;
                usedRange = usedRange.Offset[1, 0].Resize[10]; //usedRange.Rows.Count
                for (int i = 1; i < usedRange.Rows.Count; i++)
                {
                    Range rowRange = usedRange.Rows[i];
                    QRcodesHelper.GetQR(ws, rowRange);
                }
                usedRange.Cells[1, 1].Select();

                Globals.ThisAddIn.Application.Application.DisplayAlerts = true;
                Globals.ThisAddIn.Application.Application.ScreenUpdating = true;
            }
        }
    }
}
