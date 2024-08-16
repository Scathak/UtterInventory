using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.IO;

namespace UtterInventory
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            var tablesStructure = Globals.ThisAddIn.ReadCustomXML(wb, ThisAddIn.structureXMLName);
            if (tablesStructure == null)  {
                Globals.ThisAddIn.defaultStructureToCustomXml(wb);
            }
            Globals.ThisAddIn.DeployTables(wb, ThisAddIn.topLeftCornerTableRow, ThisAddIn.topLeftCornerTableCol);
            Globals.ThisAddIn.RefreshCache(wb.Sheets[ThisAddIn.rawDataSheetName]);
            Globals.ThisAddIn.DataReplication(wb, ThisAddIn.topLeftCornerTableRow, ThisAddIn.topLeftCornerTableCol);
        }
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Browse CSV Files";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.DefaultExt = "csv";
            openFileDialog1.Filter = "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|All files (*.*)|*.*";
            string outstrings;
            var wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            var ws = wb.ActiveSheet;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string file = openFileDialog1.FileName;
                using (StreamReader sr = new StreamReader(file))
                {
                    var i = 1;
                    while (sr.Peek() >= 0 && i < 1048576)
                    {
                        outstrings = sr.ReadLine();
                        if (outstrings.Length > 0)
                        {
                            var arrayOfRow = outstrings.Split(',', ';');
                            var currentRow = ws.Range[ws.Cells[1 + i, 1], ws.Cells[(1 + i), (1 + arrayOfRow.Length)]];
                            currentRow.Value2 = arrayOfRow;
                            i++;
                        }
                    }
                }
            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            var qrHelper = new QRcodesHelper();
            qrHelper.GenerateQRCodeForSelectedRange();
        }
    }
}
