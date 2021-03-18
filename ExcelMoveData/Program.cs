using OfficeOpenXml;
using System;
using System.IO;

namespace ExcelMoveData
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(fileName: @"C:\Users\user\Documents\SO-WWT ELCP at 8T Fish Farm -AE-SO20-2B1031.xlsx");
            using (ExcelPackage excel = new ExcelPackage())
            {
                
                // Create worksheets
                excel.Workbook.Worksheets.Add("WWT ELCP");
                var worksheet = excel.Workbook.Worksheets["WWT ELCP"];
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
                excel.SaveAs(file);
            }
        }
    }
}
