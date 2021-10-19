using NPOI;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WorkingHelper.Models;

namespace WorkingHelper.ExcelHandler
{
    class RetestSheetExcelOperator
    {
        public string _filePath { get; set; }
        private int GCindex = 3;
        private int FFindex = 4;
        private int GTindex = 5;
        private int GT2index = 6;
        XSSFWorkbook wb = null;

        public enum SheetEnum
        {
            yieldSheet = 1,
            retestSheet = 2,
            summarySheet = 3
        }

        public RetestSheetExcelOperator(string filePath)
        {
            _filePath = filePath;

            using (FileStream fs = File.Open(_filePath, FileMode.Open, FileAccess.Read))
            {
                wb = new XSSFWorkbook(fs);
            }
        }

        /// <summary>
        /// 修改Excel值
        /// </summary>
        /// <param name="sheetEnum"></param>
        /// <param name="rowindex"></param>
        /// <param name="colindex"></param>
        /// <param name="context"></param>
        public void ReviseExcelValue(SheetEnum sheetEnum, int rowindex, int colindex, int context)
        {
            ISheet sheet = wb.GetSheetAt((int)sheetEnum);
            IRow row = sheet.GetRow(rowindex);
            ICell cell = row.GetCell(colindex);
            cell.SetCellValue(context);

            using (FileStream fileStream = File.Open(_filePath, FileMode.Create, FileAccess.Write))
            {
                wb.Write(fileStream);
            }
        }

        /// <summary>
        /// 设置单元格格式
        /// </summary>
        /// <param name="sheetEnum"></param>
        /// <param name="rowindex"></param>
        /// <param name="colindex"></param>
        public void SetCellBorderStyle(SheetEnum sheetEnum, int rowindex, int colindex)
        {
            ISheet sheet = wb.GetSheetAt((int)sheetEnum);

            ICellStyle style = wb.CreateCellStyle();
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            sheet.GetRow(rowindex).CreateCell(colindex).CellStyle = style;
        }

        public void RetestSheetFilling(RowCounter rowCounter, ExcelDataFromSummaryHTMLModel excelDataModel_get, params List<RetestUnitModel>[] retestUnitModels)
        {
            ISheet sheet = wb.GetSheetAt((int)SheetEnum.retestSheet);

            if (/*(rowCounter.GCRetestCount == 0) || (rowCounter.GCRetestCount == 1)*/true)
            {
                ReviseExcelValue(SheetEnum.retestSheet, GCindex, 2, int.Parse(excelDataModel_get.YieldSheet_GC_Input));
                ReviseExcelValue(SheetEnum.retestSheet, GCindex, 3, int.Parse(excelDataModel_get.RetestSheet_GC_RetestCount));
                sheet.GetRow(GCindex).GetCell(4).SetCellFormula(String.Format("D{0:G}/C{1:G}", GCindex + 1, GCindex + 1));
            }
            else
            {
                ReviseExcelValue(SheetEnum.yieldSheet, GCindex, 2, int.Parse(excelDataModel_get.YieldSheet_GC_Input));
                ReviseExcelValue(SheetEnum.yieldSheet, GCindex, 3, int.Parse(excelDataModel_get.YieldSheet_GC_Pass));
                ReviseExcelValue(SheetEnum.yieldSheet, GCindex, 4, int.Parse(excelDataModel_get.YieldSheet_GC_Fail));

                //string a = String.Format("D{0:G}/C{1:G}", GCindex, GCindex);
                sheet.GetRow(GCindex).GetCell(5).SetCellFormula(String.Format("D{0:G}/C{1:G}", GCindex + 1, GCindex + 1));

                sheet.ShiftRows(FFindex, sheet.LastRowNum, rowCounter.GCFailCount - 1, true, false);
                FFindex += rowCounter.GCFailCount - 1;
                GTindex += rowCounter.GCFailCount - 1;
                GT2index += rowCounter.GCFailCount - 1;
                CellRangeAddress region;

                for (int i = 1; i <= rowCounter.GCFailCount - 1; i++)
                {
                    sheet.CreateRow(GCindex + i);
                }

                for (int j = 0; j < 10; j++)
                {
                    for (int i = 1; i <= rowCounter.GCFailCount - 1; i++)
                    {
                        SetCellBorderStyle(SheetEnum.yieldSheet, GCindex + i, j + 1);
                    }
                }

                for (int i = 0; i < 5; i++)
                {
                    region = new CellRangeAddress(GCindex, GCindex + rowCounter.GCFailCount - 1, i + 1, i + 1);
                    sheet.AddMergedRegion(region);
                }
            }

            using (FileStream fileStream = File.Open(_filePath, FileMode.Create, FileAccess.Write))
            {
                sheet.ForceFormulaRecalculation = true;
                wb.Write(fileStream);
            }
        }
    }
}
