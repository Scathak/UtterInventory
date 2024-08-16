using Microsoft.Office.Interop.Excel;
using QRCoder;
using System.Drawing;
using System.IO;
using System.Windows.Forms;


namespace UtterInventory
{
    internal class QRcodesHelper
    {
        public void GenerateQRCodeForSelectedRange()
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range selectedRange = Globals.ThisAddIn.Application.Selection;

            // Check if only a single row is selected
            if (selectedRange.Rows.Count == 1)
            {
                GetQR(activeSheet, selectedRange);
            }
            else
            {
                MessageBox.Show("Please select cells from a single row to generate the QR code.");
            }
        }
            
        public static void GetQR(Worksheet ws, Range usedRange)
        {
            // Concatenate all cell values in the current row
            string qrText = "";
            foreach (Range cell in usedRange.Cells)
            {
                qrText += cell.Value2?.ToString() + " ";
            }

            // Generate QR Code in SVG format
            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(qrText.Trim(), QRCodeGenerator.ECCLevel.Q);
            SvgQRCode qrCode = new SvgQRCode(qrCodeData);


            string qrCodeSvg = qrCode.GetGraphic(10);  // Adjust the size as needed
            Bitmap bitmap;
            Range nextCell = usedRange.Offset[0, usedRange.Columns.Count];
            using (var ms = new MemoryStream())
            {
                byte[] svgBytes = System.Text.Encoding.UTF8.GetBytes(qrCodeSvg);
                ms.Write(svgBytes, 0, svgBytes.Length);
                ms.Position = 0;

                var svgDocument = Svg.SvgDocument.Open<Svg.SvgDocument>(ms);
                bitmap = svgDocument.Draw();
            }

            // Set the bitmap to the clipboard
            Clipboard.SetImage(bitmap);
            nextCell.Select();
            ws.Paste();
            Pictures pictures = ws.Pictures();
            pictures.Left = nextCell.Left;
            pictures.Top = nextCell.Top;
            pictures.Width = nextCell.Width;
            pictures.Height = nextCell.Height;
        }
    }
}
