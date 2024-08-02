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
                var occupiedRange = wsForCopy.ListObjects[wsForCopy.Name].Range;
                occupiedRange.Offset[1, 0].Resize[columnHeigh-1].Cells.Value2 = cacheToCopy;
                SelectOneCell(wsForCopy, occupiedRange,2,1);
            }
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
    }
}
