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

            worksheet.Cells["B2"].Value = "A";
            worksheet1.Cells["B14"].Copy(worksheet.Cells["C3"]);
            worksheet1.Cells["B15"].Copy(worksheet.Cells["D3"]);
            worksheet1.Cells["C15"].Copy(worksheet.Cells["E3"]);
            worksheet1.Cells["D15"].Copy(worksheet.Cells["F3"]);
            worksheet1.Cells["E15"].Copy(worksheet.Cells["G3"]);
            worksheet1.Cells["F15"].Copy(worksheet.Cells["H3"]);
            worksheet1.Cells["I15"].Copy(worksheet.Cells["M3"]);
            worksheet1.Cells["J15"].Copy(worksheet.Cells["N3"]);

            worksheet.Column(4).Width = 60;

            package1.SaveAs(file1);

            package.SaveAs(file);
            
        }
    }
}
