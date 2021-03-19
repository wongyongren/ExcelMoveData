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
            


            CopyData();
            //CreateHead();


        }
        private static void CreateHead()
        {
            /*            var file = new FileInfo(fileName: @"C:\Users\user\Documents\Excel\CopyExcel.xlsx");
                        var file1 = new FileInfo(fileName: @"C:\Users\user\Documents\Excel\8T WWT ELCP Panel Quotation @ 8T Fish Farm.xlsx");
                        using var package = new ExcelPackage(file);
                        using var package1 = new ExcelPackage(file1);


                        // Create worksheets
                            package.Workbook.Worksheets.Add("WWT ELCP");
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
                        package.SaveAs(file);*/

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


            worksheet1.Cells["B15"].Copy(worksheet.Cells["D2"]);

            package1.SaveAs(file1);
            package.SaveAs(file);
            
        }
    }
}
