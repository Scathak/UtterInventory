using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace UtterInventory
{
    public partial class ThisAddIn
    {
        public const string rawDataSheetName = "_rawData";
        public const string structureXMLName = "structure";
        public const int topLeftCornerTableRow = 7;
        public const int topLeftCornerTableCol = 1;
        public object[,] RawDataCache;
        public int totalNumberOfStyles = 0;
        private int currentStyle = 24;
        public Dictionary<string, int> stylesForTables = new Dictionary<string, int>();

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
            { "#d8374", "Exploitation period" },
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
            { "@49dfb", rawDataSheetName },
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
            RawDataCache = OccupiedDataRange.Cells.Value2 as object[,];
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
        public void DeployTables(Workbook wb, int row, int col)
        {
            foreach (var table in GetAllTablesNames())
            {
                if (!workSheetExist(wb, table.Value))
                {
                    CreateTable(wb, table.Value, GetTablesStructure()[table.Key], row, col);
                }
            }
        }
        public bool workSheetExist(Workbook wb, string sheetName)
        {
            foreach(Worksheet worksheet in wb.Sheets)
            {
                if (worksheet.Name == sheetName) { return true; }
            }
            return false;
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

                if (!stylesForTables.ContainsKey(tableName))
                {
                    stylesForTables.Add(tableName, currentStyle++);
                    if (currentStyle > totalNumberOfStyles) currentStyle = 0;
                }
            }
            var selectedValues = GetAllStrings()
               .Where(x => columnsKeys.Contains(x.Key))
               .Select(x => x.Value)
               .ToArray();
            if (tableName == rawDataSheetName) 
            {
                deployRawDataHeaders(GetAllStrings().Keys.ToArray(), ws);
            }
            else
            {
                deployHeaders(tableName, selectedValues, row, col, ws);
            }
        }
        public void deployTablesFromXml(Workbook wb, Structure structure, int row, int col)
        {
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
            }
        }
        public void applyTableStyles(Worksheet ws, Range tableRawData)
        {
            tableRawData.Select();
            Globals.ThisAddIn.Application.CutCopyMode = XlCutCopyMode.xlCut;
            if(!objectExist(ws, ws.Name)) {
                ws.ListObjects.Add(XlListObjectSourceType.xlSrcRange, tableRawData, Type.Missing, XlYesNoGuess.xlYes).Name = ws.Name;
                ws.ListObjects[ws.Name].TableStyle = Globals.ThisAddIn.Application.ActiveWorkbook.TableStyles.Item(stylesForTables[ws.Name]);
                ws.Tab.Color = XlRgbColor.rgbBlue;
            }
        }
        public void selectRow(Worksheet ws, int firstCornerRow, int firstCornerColumn, int tableWidth, int rowToSelect)
        {
            ws.Range[ws.Cells[rowToSelect + firstCornerRow, firstCornerColumn], ws.Cells[rowToSelect + firstCornerRow, tableWidth + firstCornerColumn]].Select();
        }
        public bool objectExist(Worksheet ws, string objName)
        {
            foreach (ListObject currentObject in ws.ListObjects)
            {
                if(currentObject.Name == objName) return true;
            }
            return false;
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
                var tableWidth = headings.Length - 1;
                var columnHeigh = RawDataCache.GetLength(0);
                Range tableRawData = ws.Range[ws.Cells[row, col], ws.Cells[(row + columnHeigh), (col + tableWidth)]];
                applyTableStyles(ws, tableRawData);
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
