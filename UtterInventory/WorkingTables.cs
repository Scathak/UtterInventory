using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;


namespace UtterInventory
{
    public partial class ThisAddIn
    {
        public object[,] RawDataCache;

        public Dictionary<string, string[]> TablesStructure = new Dictionary<string, string[]>()
        {
            { "@abb37", new string[] { "#df77e", "#a51aa", "#475bd", "#3fb52", "#7da2c", "#27843", "#64c4a", "#a7a55", "#c5bf4", "#81a8f" } },
            { "@986cd", new string[] { "#df77e", "#a51aa", "#14ddd", "#a6527", "#d4d25", "#a47a5", "#cd5d4" } },
            { "@feab5", new string[] { "#df77e", "#a51aa", "#475bd", "#3fb52", "#7da2c", "#27843", "#64c4a", "#a7a55", "#9258d", "#46a88", "#c5bf4", "#54c01", "#81a8f", "#5d5ae", "#4f29f", "#304b2", "#a47a5", "#cd5d4" }},
            { "@ca35f", new string[] { "#df77e", "#a51aa", "#475bd", "#14ddd", "#a6527" }},
            { "@49dfb", new string[] { "#df77e", "#a51aa", "#475bd", "#3fb52", "#7da2c", "#27843", "#64c4a", "#a7a55", "#d8374", "#c11f8", "#1c53e", "#14ddd", "#a6527", "#9258d", "#46a88", "#c5bf4", "#54c01",  "#d4d25",  "#81a8f",  "#5d5ae", "#4f29f", "#304b2", "#a47a5", "#cd5d4"} }
        };

        private Dictionary<string, string> AllKeys = new Dictionary<string, string>()
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
        private Dictionary<string, string> AllTablesNames = new Dictionary<string, string>()
        {
            { "@abb37", "Inventory List" },
            { "@986cd", "Financial Balance" },
            { "@feab5", "Movement of Inventories" },
            { "@ca35f", "Balance Worksheet" },
            { "@49dfb", "_rawData" },
        };
        private Dictionary<string, string> TablesToCopyOn = new Dictionary<string, string>()
        {
            { "@abb37", "Inventory List" },
            { "@986cd", "Financial Balance" },
            { "@feab5", "Movement of Inventories" },
            { "@ca35f", "Balance Worksheet" }
        };
        private string[] specificStrings = new string[] { "Date of Inquiry", "Expiry Date", "Department", "Category", "Quantity in Reorder" };
        public void RefreshCache(Worksheet ws)
        {
            Range OccupiedDataRange = ws.Range[ws.Cells[1, 1], GetOccupiedCells(ws)];
            RawDataCache = OccupiedDataRange.Cells.Value2;
        }
        public Dictionary<string, string> GetAllTablesNames()
        {
            return AllTablesNames;
        }
        public Dictionary<string, string> GetTablesToCopyOn()
        {
            return TablesToCopyOn;
        }
        public Dictionary<string, string[]> GetTablesStructure()
        {
            return TablesStructure;
        }
        public Dictionary<string, string> GetAllStrings()
        {
            return AllKeys;
        }
        string[] GetTableSpecificStrings()
        {
            return specificStrings;
        }
        public void DeployTables(int row, int col)
        {
            var wb = Application.ActiveWorkbook;
            Worksheet ws = wb.ActiveSheet;
            foreach (var table in GetAllTablesNames())
            {
                CreateTable(wb, table.Value, GetTablesStructure()[table.Key], row, col);
            }
        }
        public void CreateTable(Workbook wb, string tableName, string[] columnsKeys, int row, int col)
        {
            Worksheet ws = null;
            foreach (Worksheet sheet in wb.Worksheets)
            {
                if (sheet.Name == tableName)
                {
                    ws = sheet;
                    break;
                }
            }
            if (ws == null)
            {
                ws = wb.Worksheets.Add(Type: XlSheetType.xlWorksheet);
                ws.Name = tableName;
            }
            var selectedValues = GetAllStrings()
               .Where(x => columnsKeys.Contains(x.Key))
               .Select(x => x.Value)
               .ToArray();
            if (tableName == "_rawData") 
            {
                deployRawDataHeaders(GetAllStrings().Keys.ToArray(), ws);
            }
            else
            {
                deployHeaders(tableName, selectedValues, row, col, ws);
            }
        }
        public void deployTablesFromXml(Structure structure, int row, int col)
        {
            var wb = Application.ActiveWorkbook;
            foreach (var table in structure.Tables)
            {
                var ws = (Worksheet)wb.Worksheets.Add(Type: XlSheetType.xlWorksheet);
                var selectedColNames = GetAllStrings()
                   .Where(x => table.ColumnsNames.Contains(x.Key))
                   .Select(x => x.Value)
                   .ToArray();
                var selectedTableName = GetAllTablesNames()[table.TableName];
                deployHeaders(selectedTableName, selectedColNames, row, col, ws);
            };
        }
        public void deployRawDataHeaders(string[] headings, Worksheet ws)
        {
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
                RefreshCache(ws);
            }
        }
        public void deployHeaders(string WSName, string[] headings, int row, int col, Worksheet ws)
        { 
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
    }
}
