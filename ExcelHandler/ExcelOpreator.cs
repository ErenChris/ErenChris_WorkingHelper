using NPOI;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.Util;
using System.IO;

namespace WorkingHelper.ExcelHandler
{
    class ExcelOpreator
    {
        private string _filePath { get; set; }

        public enum SheetEnum
        {
            yieldSheet = 1,
            retestSheet = 2,
            summarySheet = 3
        }

        public ExcelOpreator(string filePath)
        {
            _filePath = filePath;
        }

        public void ReviseExcelValue(SheetEnum sheetEnum, int rowindex, int colindex, string context)
        {
            XSSFWorkbook wb = null;

            using (FileStream fs = File.Open(_filePath, FileMode.Open, FileAccess.Read))
            {
                wb = new XSSFWorkbook(fs);
            }

            ISheet sheet = wb.GetSheetAt((int)sheetEnum);
            IRow row = sheet.GetRow(rowindex - 1);
            ICell cell = row.GetCell(colindex - 1);
            cell.SetCellValue(context);

            using (FileStream fileStream = File.Open(_filePath, FileMode.Create, FileAccess.Write))
            {
                wb.Write(fileStream);
            }
        }

        public string GetExcelValueFromSingleCell(SheetEnum sheetEnum, int rowindex, int colindex)
        {
            XSSFWorkbook wb = null;

            using (FileStream fs = File.Open(_filePath, FileMode.Open, FileAccess.Read))
            {
                wb = new XSSFWorkbook(fs);
            }

            ISheet sheet = wb.GetSheetAt((int)sheetEnum);
            string result = sheet.GetRow(rowindex - 1).GetCell(colindex - 1).ToString();

            return result;
        }

        public string GetExcelValueFromSingleCell(ICell cell)
        {
            string result = cell.ToString();

            return result;
        }

        public string GetExcelValueFromMergeCells(SheetEnum sheetEnum, int rowindex, int colindex)
        {
            XSSFWorkbook wb = null;
            string result = null;

            using (FileStream fs = File.Open(_filePath, FileMode.Open, FileAccess.Read))
            {
                wb = new XSSFWorkbook(fs);
            }

            ISheet sheet = wb.GetSheetAt((int)sheetEnum);
            ICell cell = sheet.GetRow(rowindex - 1).GetCell(colindex - 1);
            if (cell.IsMergedCell)
            {
                string temp = GetExcelValueFromSingleCell(cell);
                if (string.IsNullOrEmpty(temp))
                {
                    for (int i = 0; i < rowindex; i++)
                    {
                        cell = sheet.GetRow(rowindex - (2 + i)).GetCell(colindex - 1);
                        temp = GetExcelValueFromSingleCell(cell);
                        if (!string.IsNullOrEmpty(temp))
                        {
                            result = temp;

                            return result;
                        }
                    }
                }
                else
                {
                    result = temp;
                }
            }

            return result;
        }
    }
}
