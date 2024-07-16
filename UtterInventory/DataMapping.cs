using System;
using Microsoft.Office.Interop.Excel;


namespace UtterInventory
{
    public partial class ThisAddIn
    {
        public void DataReplication(Workbook wb, int row, int col)
        {
            Range lastCell = GetOccupiedCells(wb.ActiveSheet);
            //Range rrr = wb.Application.get_Range("A1", lastCell);
            var tableRows = lastCell.Rows.Count - row;
            var tableColumns = lastCell.Columns.Count - col;
            var rawDataSheet = wb.Sheets["_rawData"];

            Range OccupiedDataRange = rawDataSheet.Range[rawDataSheet.Cells[row, col], GetOccupiedCells(rawDataSheet)];
            //rawDataSheet.Activate();
            //OccupiedDataRange.Select();
            //var aa = rawDataSheet.Range.Find(InventoryKeys[0], LookAt: XlLookAt.xlWhole);
            //var C = new object[1000,100] as System.__ComObject;
            /*C = rawDataSheet.Cells;
            var A = rawDataSheet.Cells[1,1].Value2;
            var B = A.GetType();
            C[1, 1] = A;
            var t = C.GetType();*/
        }
        public Range GetOccupiedCells(Worksheet ws)
        {
            return ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
        }
    }
}
