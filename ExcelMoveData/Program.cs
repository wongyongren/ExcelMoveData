using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Drawing;
using System.IO;

namespace ExcelMoveData
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            CreateHead();
            TestData();
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

            //Create Header
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
            //Adjust Columns Width
            var range = worksheet.Cells["A1:N1"];
            range.AutoFitColumns();
            package.SaveAs(file);
        }

        private static void TestData()
        {
            //Create and Link the New File
            var file = new FileInfo(fileName: @"C:\Users\user\Documents\Excel\CopyExcel.xlsx");
            //The Original File
            var file1 = new FileInfo(fileName: @"C:\Users\user\Documents\Excel\8T WWT ELCP Panel Quotation @ 8T Fish Farm.xlsx");
            using var package = new ExcelPackage(file);
            using var package1 = new ExcelPackage(file1);
            // from which worksheet and worksheet name
            var worksheet = package.Workbook.Worksheets["WWT ELCP"];
            var worksheet1 = package1.Workbook.Worksheets["WWT ELCP"];

            //define for Row,value and column value
            var row1 = 14;
            double value = 1;
            char column = 'A';

            // For Loop
            for (var row = 2; row < 1000; row++)
            {
                // if original file column get null value or get column value of alphabet
                if ((worksheet1.Cells[$"A{row1}"].Value ?? "").ToString() != "" && worksheet1.Cells[$"A{row1}"].Value.ToString().Equals(column.ToString()))
                {
                    //copy from original file to new file
                    worksheet1.Cells[$"A{row1}"].Copy(worksheet.Cells[$"B{row}"]);
                    worksheet.Cells[$"A{row}"].Value = "M";
                    worksheet1.Cells[$"B{row1}"].Copy(worksheet.Cells[$"C{row}"]);
                    worksheet1.Cells[$"C{row1}"].Copy(worksheet.Cells[$"D{row}"]);
                    worksheet1.Cells[$"D{row1}"].Copy(worksheet.Cells[$"E{row}"]);
                    worksheet1.Cells[$"E{row1}"].Copy(worksheet.Cells[$"F{row}"]);
                    worksheet1.Cells[$"F{row1}"].Copy(worksheet.Cells[$"G{row}"]);
                    worksheet1.Cells[$"G{row1}"].Copy(worksheet.Cells[$"H{row}"]);
                    worksheet1.Cells[$"J{row1}"].Copy(worksheet.Cells[$"M{row}"]);
                    worksheet1.Cells[$"K{row1}"].Copy(worksheet.Cells[$"N{row}"]);
                    //row1 is using for original file because both row is different row
                    row1++;
                }
                // if original file column get value of 1 ,1.1
                else if ((worksheet1.Cells[$"A{row1}"].Value ?? "").ToString() != "" && worksheet1.Cells[$"A{row1}"].Value.ToString().Equals(value.ToString()))
                {
                    worksheet1.Cells[$"A{row1}"].Copy(worksheet.Cells[$"B{row}"]);
                    worksheet.Cells[$"A{row}"].Value = "S";
                    worksheet.Cells[$"B{row}"].Value = $"{column}." + worksheet.Cells[$"B{row}"].Value;
                    worksheet1.Cells[$"B{row1}"].Copy(worksheet.Cells[$"C{row}"]);
                    worksheet1.Cells[$"C{row1}"].Copy(worksheet.Cells[$"D{row}"]);
                    worksheet1.Cells[$"D{row1}"].Copy(worksheet.Cells[$"E{row}"]);
                    worksheet1.Cells[$"E{row1}"].Copy(worksheet.Cells[$"F{row}"]);
                    worksheet1.Cells[$"F{row1}"].Copy(worksheet.Cells[$"G{row}"]);
                    worksheet1.Cells[$"G{row1}"].Copy(worksheet.Cells[$"H{row}"]);
                    worksheet1.Cells[$"J{row1}"].Copy(worksheet.Cells[$"M{row}"]);
                    worksheet1.Cells[$"K{row1}"].Copy(worksheet.Cells[$"N{row}"]);
                    row1++;
                    //do addition to get 1.1,1.2
                    value = Math.Round(value, 1) + 0.1;
                    value = Math.Round(value, 1);
                }
                //if original file get null value or column equal to setting column,it will break and stop the for loop
                else if ((worksheet1.Cells[$"A{row1}"].Value ?? "").ToString() == "" || column == 'E')
                {
                    break;
                }
                else
                {
                    //when column value is not equal to column value, it will change to next alphabet value and set value back to 1
                    ++column;
                    System.Console.WriteLine(column);
                    value = 1;
                }
            }
            //set row style
            for (var row = 2; row < 1000; row++)
            {
                worksheet.Cells[$"B{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[$"E{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"F{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"G{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"H{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"M{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"N{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            //set column style
            for (var i = 9; i < 15; i++)
            {
                worksheet.Column(i).Style.Font.Color.SetColor(Color.Red);
                worksheet.Column(i).Style.Numberformat.Format = "General";
            }

            //set entire worksheet style
            worksheet.Column(4).Width = 60;
            worksheet.Column(3).Width = 11;
            worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
            worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            worksheet.Cells.Style.Font.Size = 11;
            worksheet.Cells.Style.Font.Name = "Arial";
            worksheet.Cells.Style.Font.Bold = false;
            worksheet.View.FreezePanes(2, 1);
            package.SaveAs(file);
        }
    }
}