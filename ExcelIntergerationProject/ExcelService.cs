using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace ExcelIntergerationProject
{
    internal class ExcelService
    {
        public void Export(string Title, string ColumnHead01, string ColumnHead02, Dictionary<string, string> dic)
        {
            string filePath = SaveFile();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(Title);

                // Add header row
                worksheet.Cells[1, 1].Value = ColumnHead01;
                worksheet.Cells[1, 2].Value = ColumnHead02;


                int row = 2;

                // Iterate through each layer group
                foreach (var item in dic)
                {
                    worksheet.Cells[row, 1].Value = item.Key;
                    worksheet.Cells[row, 2].Value = item.Value;
                    row++;
                }

                // Adjust column widths to fit content
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                package.SaveAs(new FileInfo(filePath));
            }

            // Open the file
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = filePath,
                UseShellExecute = true
            });
        }

        public void Import(string filePath)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(Title);

                // Add header row
             var item01 =   worksheet.Cells[1, 1].Value;


                int row = 2;

                // Iterate through each layer group
                foreach (var item in dic)
                {
                    worksheet.Cells[row, 1].Value = item.Key;
                    worksheet.Cells[row, 2].Value = item.Value;
                    row++;
                }

                // Adjust column widths to fit content
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                package.SaveAs(new FileInfo(filePath));
            }

            // Open the file
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = filePath,
                UseShellExecute = true
            });
        }

        private string SaveFile()
        {
            string filePath = null;
            SaveFileDialog saveFileDialog = new SaveFileDialog()
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Title = "Select a Location to Save the Excel File",
                FileName = "RoomArea.xlsx"
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Get the file path from the dialog
                filePath = saveFileDialog.FileName;
            }
            return filePath;
        }

    }
}
