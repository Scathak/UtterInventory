using Microsoft.Office.Interop.Excel;
using System.Linq;
using System;

namespace UtterInventory
{
    public partial class ThisAddIn
    {
        public void DataReplication(Workbook wb, int row, int col)
        {
            Worksheet rawDataSheet = wb.Sheets["_rawData"];
            rawDataSheet.Activate();
            Range lastCell = GetOccupiedCells(wb.ActiveSheet);

            RefreshCache(rawDataSheet);

            Range OccupiedDataRange = rawDataSheet.Range[rawDataSheet.Cells[1, 1], GetOccupiedCells(rawDataSheet)];

            foreach (var table in TablesStructure)
            {
                Worksheet wsForCopy = null;
                if (GetTablesToCopyOn().ContainsKey(table.Key))
                {
                    wsForCopy = wb.Worksheets[GetAllTablesNames()[table.Key]];
                }else break;

                var j = 0;
                var colNames = table.Value;
                foreach (var column in colNames)
                {
                    var tableWidth = RawDataCache.GetLength(1);

                    var tableHeaders = Enumerable.Range(1, RawDataCache.GetLength(1))
                                .Select(x => RawDataCache[1, x])
                                .ToArray();
                    var IndexOccurence = Array.FindIndex(tableHeaders, w => w.Equals(column));
                    if (IndexOccurence >= 0)
                    {
                        var tableColumn = Enumerable.Range(1 + 1, RawDataCache.GetLength(0) - 1)
                            .Select(x => RawDataCache[x, IndexOccurence + 1])
                            .ToArray();
                        wsForCopy.Range[wsForCopy.Cells[row + 1, col + j], wsForCopy.Cells[row + tableColumn.Length, col + j]].Cells.Value2 = tableColumn;
                        j++;
                    }
                }
            }
        }
        public Range GetOccupiedCells(Worksheet ws)
        {
            return ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
        }
    }
}
