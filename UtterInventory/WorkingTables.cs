using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace UtterInventory
{
    public partial class ThisAddIn
    {
        public void DeployTables(int row, int col)
        {
            if (OnStartEmpty)
            {
                var wb = Application.ActiveWorkbook;
                Worksheet ws;
                var isPresent = false;
                foreach(Worksheet sheet in wb.Worksheets)
                {
                    if (sheet.Name == "_rawData")
                        isPresent = true;
                }
                if (!isPresent)
                {
                    ws = (Excel.Worksheet)wb.Worksheets.Add(Type: XlSheetType.xlWorksheet);

                    deployRawDataHeaders(GetAllStrings().Keys.ToArray(), ws);
                }

                ws = (Excel.Worksheet)wb.Worksheets.Add(Type: XlSheetType.xlWorksheet);
                var  selectedValues = GetAllStrings()
                   .Where(x => InventoryListKeys().Contains(x.Key))
                   .Select(x => x.Value)
                   .ToArray();
                deployHeaders("Inventory List", selectedValues, row, col, ws);

                ws = (Excel.Worksheet)wb.Worksheets.Add(Type: XlSheetType.xlWorksheet);
                selectedValues = GetAllStrings()
                   .Where(x => FinancialBalanceKeys().Contains(x.Key))
                   .Select(x => x.Value)
                   .ToArray();
                deployHeaders("Financial Balance", selectedValues, row, col, ws);

                ws = (Excel.Worksheet)wb.Worksheets.Add(Type: XlSheetType.xlWorksheet);
                selectedValues = GetAllStrings()
                   .Where(x => MovementInventoriesKeys().Contains(x.Key))
                   .Select(x => x.Value)
                   .ToArray();
                deployHeaders("Movement of Inventories", selectedValues, row, col, ws);

                ws = (Excel.Worksheet)wb.Worksheets.Add(Type: XlSheetType.xlWorksheet);
                selectedValues = GetAllStrings()
                   .Where(x => BalanceSheetKeys().Contains(x.Key))
                   .Select(x => x.Value)
                   .ToArray();
                deployHeaders("Balance Worksheet", selectedValues, row, col, ws);
            }
        }
        public void deployRawDataHeaders(string[] headings, Excel.Worksheet ws)
        {
            ws.Name = "_rawData";
            ws.Columns.AutoFit();
            ws.Cells.ColumnWidth = 14;
            ws.Cells.Font.Size = 8;
            if (headings != null && headings.Length > 0)
            {
                var tableWidth = headings?.Length - 1;
                var i = 0;
                foreach (var columnName in headings)
                {
                    ws.Cells[1, i + 1].Value = columnName;
                    i++;
                }
            }
        }
        public void deployHeaders(string WSName, string[] headings, int row, int col, Excel.Worksheet ws)
        {
            ws.Name = WSName;
            ws.Columns.AutoFit();
            ws.Cells.ColumnWidth = 14;
            ws.Cells.Font.Size = 8;
            ws.EnableOutlining = false;
            Application.ActiveWindow.DisplayGridlines = false;
            if (headings != null && headings.Length > 0)
            {
                var tableWidth = headings?.Length - 1;
                Range tableRawData = ws.Range[ws.Cells[row, col], ws.Cells[(row + 1), (col + tableWidth)]];
                Range headingsRange = ws.Range[ws.Cells[row, col], ws.Cells[(row), (col + tableWidth)]];
                Range dataOnly = ws.Range[ws.Cells[row + 1, col], ws.Cells[(row + 2), (col + tableWidth)]];
                ws.Names.Add("Table", tableRawData);
                ws.Names.Add("Headings", headingsRange);
                ws.Names.Add("Data", dataOnly);
                headingsRange.Font.Bold = true;
                headingsRange.Interior.Color = Color.LightYellow;
                tableRawData.Borders.Color = Color.LightGray;
                dataOnly.Interior.Color = Color.FromArgb(0, 243, 243, 243);
                var i = 0;
                foreach (var columnName in headings)
                {
                    ws.Cells[row, i + col].Value = columnName;
                    i++;
                }
                ws.Cells[1, 1].Value2 = WSName;
                var tableSpecStrings = GetTableSpecificStrings();
                i = 2;
                foreach (var item in tableSpecStrings)
                {
                    ws.Cells[i++, 1].Value2 = item;
                }
            }
        }
        public Dictionary<string,string> GetAllStrings()
        {
            return new Dictionary<string, string>
            {
                { "#df77e", "Inventary Number" },
                { "#a51aa", "Name" },
                { "#475bd", "Description" },
                { "#3fb52", "ProductID" },
                { "#7da2c", "Model" },
                { "#27843", "Color" },
                { "#64c4a", "Size" },
                { "#a7a55", "Weight" },
                { "#d8374", "Expiry Date" },
                { "#c11f8", "Category" },
                { "#1c53e", "Department" },
                { "#14ddd", "Unit" },
                { "#a6527", "Unit price" },
                { "#9258d", "Stock In" },
                { "#46a88", "Stock Out" },
                { "#c5bf4", "Quantity in Stock" },
                { "#54c01", "Stock Check" },
                { "#d4d25", "Inventary Value" },
                { "#81a8f", "Remark" },
                { "#5d5ae", "Reorder level" },
                { "#4f29f", "Reorder Time" },
                { "#304b2", "Quantity in Reorder" },
                { "#a47a5", "Date" },
                { "#cd5d4", "Time" }
            };
        }
        string[] GetInventoryListStrings()
        {
            string[] strings = { "Inventary Number", "Name", "Description", "ProductID", "Model", "Color", "Size", "Weight", "Quantity in Stock", "Remark" };
            return strings;
        }
        string[] InventoryListKeys()
        {
            return new string[] { "#df77e", "#a51aa", "#475bd", "#3fb52", "#7da2c", "#27843", "#64c4a", "#a7a55", "#c5bf4", "#81a8f" };
        }
        string[] GetFinancialBalanceStrings()
        {
            string[] strings = { "Inventary Number", "Name", "Unit", "Unit price", "Inventary Value", "Date", "Time" };
            return strings;
        }
        string[] FinancialBalanceKeys()
        {
            return new string[] { "#df77e", "#a51aa", "#14ddd", "#a6527", "#d4d25", "#a47a5", "#cd5d4" };
        }
        string[] GetMovementInventoriesStrings()
        {
            string[] strings = { "Inventary Number", "Name", "Description", "ProductID", "Model", "Color", "Size", "Weight", "Stock In", "Stock Out", "Quantity in Stock", "Stock Check", "Remark", "Reorder level", "Reorder Time", "Quantity in Reorder", "Date", "Time" };
            return strings;
        }
        string[] MovementInventoriesKeys()
        {
            return new string[] { "#df77e", "#a51aa", "#475bd", "#3fb52", "#7da2c", "#27843", "#64c4a", "#a7a55", "#9258d", "#46a88", "#c5bf4", "#54c01", "#81a8f", "#5d5ae", "#4f29f", "#304b2", "#c5bf4", "#81a8f", "#a47a5", "#cd5d4" };
        }
        string[] GetBalanceSheetStrings()
        {
            string[] strings = { "Inventary Number", "Name", "Description", "Unit", "Unit price" };
            return strings;
        }
        string[] BalanceSheetKeys()
        {
            return new string[] { "#df77e", "#a51aa", "#475bd", "#14ddd", "#a6527" };
        }
        string[] GetTableSpecificStrings()
        {
            string[] strings = { "Date of Inquiry", "Expiry Date", "Department", "Category", "Quantity in Reorder" };
            return strings;
        }
    }
}
