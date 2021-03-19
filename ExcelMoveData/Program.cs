using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;

namespace ExcelMoveData
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //CreateHead();
            CopyData();
        }

        private static void CreateHead()
        {
            var file = new FileInfo(fileName: @"C:\Users\user\Documents\Excel\CopyExcel.xlsx");
            var file1 = new FileInfo(fileName: @"C:\Users\user\Documents\Excel\8T WWT ELCP Panel Quotation @ 8T Fish Farm.xlsx");
            using var package = new ExcelPackage(file);
            using var package1 = new ExcelPackage(file1);

            // Create worksheets
            //package.Workbook.Worksheets.Add("WWT ELCP");
            var worksheet = package.Workbook.Worksheets["WWT ELCP"];
            //head
            worksheet.Cells["A1"].Value = "PhaseType";
            worksheet.Cells["B1"].Value = "Item";
            worksheet.Cells["C1"].Value = "Description";
            worksheet.Cells["D1"].Value = "Remarks";
            worksheet.Cells["E1"].Value = "Qty";
            worksheet.Cells["F1"].Value = "UOM";
            worksheet.Cells["G1"].Value = "UniRate";
            worksheet.Cells["H1"].Value = "Amount";
            worksheet.Cells["I1"].Value = "costcode";
            worksheet.Cells["J1"].Value = "CstStockCode";
            worksheet.Cells["K1"].Value = "CstStockDescription";
            worksheet.Cells["L1"].Value = "CstStkQty";
            worksheet.Cells["M1"].Value = "CstStockUnitRate";
            worksheet.Cells["N1"].Value = "CostAmount";

            var range = worksheet.Cells["A1:N1"];
            range.AutoFitColumns();
            package.SaveAs(file);
        }

        private static void CopyData()
        {
            var file = new FileInfo(fileName: @"C:\Users\user\Documents\Excel\CopyExcel.xlsx");
            var file1 = new FileInfo(fileName: @"C:\Users\user\Documents\Excel\8T WWT ELCP Panel Quotation @ 8T Fish Farm.xlsx");
            using var package = new ExcelPackage(file);
            using var package1 = new ExcelPackage(file1);
            var worksheet = package.Workbook.Worksheets["WWT ELCP"];
            var worksheet1 = package1.Workbook.Worksheets["WWT ELCP"];

            // headers

            worksheet.Cells["B2"].Value = "A";
            var row1 = 14;
            for (var row = 3; row < 58; row++)
            {
                worksheet1.Cells[$"B{row1}"].Copy(worksheet.Cells[$"C{row}"]);
                worksheet1.Cells[$"C{row1}"].Copy(worksheet.Cells[$"D{row}"]);
                worksheet1.Cells[$"D{row1}"].Copy(worksheet.Cells[$"E{row}"]);
                worksheet1.Cells[$"E{row1}"].Copy(worksheet.Cells[$"F{row}"]);
                worksheet1.Cells[$"F{row1}"].Copy(worksheet.Cells[$"G{row}"]);
                worksheet1.Cells[$"G{row1}"].Copy(worksheet.Cells[$"H{row}"]);
                worksheet1.Cells[$"J{row1}"].Copy(worksheet.Cells[$"M{row}"]);
                worksheet1.Cells[$"K{row1}"].Copy(worksheet.Cells[$"N{row}"]);
                worksheet.Column(row).AutoFit();
                row1++;
            }
            worksheet.Column(4).Width = 60;
            worksheet.Column(3).Width = 13;

            for (var row = 5; row < 58; row++)
            {
                worksheet.Cells[$"E{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"F{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"G{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"H{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"M{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"N{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            for (var i = 9; i < 15; i++)
            {
                worksheet.Column(i).Style.Font.Color.SetColor(Color.Red);
                worksheet1.Column(i).Style.Numberformat.Format = "General";
            }

            package1.SaveAs(file1);
            package.SaveAs(file);
        }
    }
}