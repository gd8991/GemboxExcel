using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using GemBox.Spreadsheet;

namespace GemboxExcel.Helpers
{
    public static class GemboxHelper
    {
        private static DataTable _sampleData;
        private static string _fileNameWithExtension;


        public static string GenerateFile(int noOfRows, string format)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            var fileName = "GemboxDemo";
            _fileNameWithExtension = $"{fileName}.{format}";
            _sampleData = SampleDataHelper.GenerateSampleData(noOfRows);
            var sheet = CreateSheet();
            switch (format.ToLower())
            {
                case "xlsx":
                    SaveExcelFile(sheet);
                    break;
                case "csv":
                    SaveCsFilev(sheet);
                    break;
                case "pdf":
                    SavePdfFile(sheet);
                    break;
            }

            //CreateExcel(fileName);
            return _fileNameWithExtension;
        }

        private static ExcelWorksheet CreateSheet()
        {
            ExcelFile wkBook = new ExcelFile();
            ExcelWorksheet sheet = wkBook.Worksheets.Add("Transaction Details");
            sheet.InsertDataTable(_sampleData, new InsertDataTableOptions
            {
                ColumnHeaders = true,
                StartRow = 1,
                StartColumn = 1
            });
            return sheet;
        }

        private static void SaveCsFilev(ExcelWorksheet sheet)
        {
            var workbook = sheet.Parent;
            workbook.Save(_fileNameWithExtension, new CsvSaveOptions(CsvType.CommaDelimited));
        }

        private static void SaveExcelFile(ExcelWorksheet sheet)
        {
            var workbook = sheet.Parent;
            workbook.Save(_fileNameWithExtension);
        }

        private static void SavePdfFile(ExcelWorksheet sheet)
        {
            var workbook = sheet.Parent;
            workbook.Save(_fileNameWithExtension, new PdfSaveOptions());
        }       

        private static void SetStyleToTable(ExcelWorksheet sheet)
        {
            //for (int i = 0; i <= 52; i++)
            //{
            //    for (int j = 4; j <= 5; j++)
            //    {

            //        if (sheet.Rows[i].CellList[j].Value != null && sheet.Rows[i].CellList[j].Value.ToString().Contains('-'))
            //        {
            //            sheet.Rows[i].CellList[j].Style.Font.Color = System.Drawing.Color.Red;
            //        }
            //        else if (sheet.Rows[i].CellList[j].Value is string && Convert.ToString(sheet.Rows[i].CellList[j].Value).Equals("Overdraft"))
            //        {
            //            sheet.Rows[i].CellList[j].Style.Font.Color = System.Drawing.Color.White;
            //            sheet.Rows[i].CellList[j].Style.Color = System.Drawing.Color.Red;
            //        }
            //    }
            //}
        }

        private static void CreateTableBorders(ExcelWorksheet sheet)
        {
           //var table = sheet.Tables.Add()
        }

        private static void CreatePivot(ExcelFile wkBook)
        {
            //CellRange range = wkBook.Worksheets[0].Range["A1:F52"];
            //Worksheet pivotSheet = wkBook.Worksheets[0];
            //pivotSheet.Name = "Pivot";
            //PivotCache cache = wkBook.PivotCaches.Add(range);
            //PivotTable pt = pivotSheet.PivotTables.Add("Pivot Table", wkBook.Worksheets[0].Range["J2"], cache);

            ////XlsPivotTable pivotTable = wkBook.Worksheets[0].PivotTables.Add("PivotTable", range, cache);
            ////pivotTable.Cache.IsRefreshOnLoad = true;

            ////var r1 = pt.PivotFields["Description"];
            ////r1.Axis = AxisTypes.Row;
            ////pt.Options.RowHeaderCaption = "Description";

            //var r2 = pt.PivotFields["Transaction Type"];
            //pt.Options.RowHeaderCaption = "Transaction Type";
            //r2.Axis = AxisTypes.Row;

            //pt.DataFields.Add(pt.PivotFields["Transaction Type"], "Sum of Transactions", SubtotalTypes.Count);

            //pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium12;

        }

        private static void GetImage(ExcelFile workbook)
        {

            //var imageStream = workbook.Worksheets[0].ToImage(0, 0, 530, 100);
            //MemoryStream ms = new MemoryStream();
            //CopyStream(imageStream, ms);
            //System.Drawing.Image img = System.Drawing.Image.FromStream(ms);

            //Bitmap bmp = img as Bitmap;

        }
        //private void CopyStream(Stream input, Stream output)
        //{
        //    byte[] buffer = new byte[16 * 1024];
        //    int read;
        //    while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
        //    {
        //        output.Write(buffer, 0, read);
        //    }
        //}

        private static void PageSetup(ExcelWorksheet worksheet)
        {
            //worksheet.PageSetup.AlignWithMargins = 1;
            //worksheet.PageSetup.ScaleWithDoc = 0; // 
        }

        private static void RefreshPivotTables(ExcelWorksheet worksheet)
        {
            // Auto-refresh
        }

        private static void PrintWorkbook(ExcelFile workbook)
        {
            // Acc to Spire.xls website, printed is supported from v7.5.3, but i've installed v11. and not getting Print document option "workbook.PrintDocument"

        }
    }
}
