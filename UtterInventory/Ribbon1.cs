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
            Globals.ThisAddIn.DeployTables(7,1);
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
            var targetRange = ws.Range[ws.Cells[1 + 1, 1], ws.Cells[(1 + 2), (1 + 24)]];
            var targetRow = ws.Range[ws.Cells[1 + 1, 1], ws.Cells[(1 + 1), (1 + 24)]];

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
                            var currentRow = ws.Range[ws.Cells[1 + i, 1], ws.Cells[(1 + i), (1 + 24)]];
                            currentRow.Value2 = outstrings.Split(',', ';');
                            i++;
                        }
                    }
                }
            }
        }
    }
}
