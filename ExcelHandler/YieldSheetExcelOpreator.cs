using NPOI;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.Util;
using System.IO;
using WorkingHelper.Models;
using System.Collections.Generic;
using System;

namespace WorkingHelper.ExcelHandler
{
    class YieldSheetExcelOpreator
    {
        private string _filePath { get; set; }
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

        public YieldSheetExcelOpreator(string filePath)
        {
            _filePath = filePath;

            using (FileStream fs = File.Open(_filePath, FileMode.Open, FileAccess.Read))
            {
                wb = new XSSFWorkbook(fs);
            }
        }

        public void ReviseExcelValue(SheetEnum sheetEnum, int rowindex, int colindex, string context)
        {
            ISheet sheet = wb.GetSheetAt((int)sheetEnum);
            IRow row = sheet.GetRow(rowindex - 1);
            ICell cell = row.GetCell(colindex - 1);
            cell.SetCellValue(context);

            using (FileStream fileStream = File.Open(_filePath, FileMode.Create, FileAccess.Write))
            {
                wb.Write(fileStream);
            }
        }

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

        public string GetExcelValueFromSingleCell(SheetEnum sheetEnum, int rowindex, int colindex)
        {
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
            string result = null;

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

        public int GetLastRowIndex(SheetEnum sheetEnum)
        {
            int result;

            ISheet sheet = wb.GetSheetAt((int)sheetEnum);

            result = sheet.LastRowNum;

            return result;
        }

        /// <summary>
        /// 填充良率列表数据
        /// </summary>
        /// <param name="rowCounter"></param>
        /// <param name="excelDataModel_get"></param>
        /// <param name="retestUnitModels"></param>
        public void YieldSheetFilling(RowCounter rowCounter, ExcelDataFromSummaryHTMLModel excelDataModel_get, params List<RetestUnitModel>[] retestUnitModels)
        {
            ISheet sheet = wb.GetSheetAt((int)SheetEnum.yieldSheet);

            //rowCounter.GCFailCount = 3; // should be deleted
            //rowCounter.FFFailCount = 2;
            //rowCounter.GTFailCount = 4;
            //rowCounter.GT2FailCount = 5;

            if ((rowCounter.GCFailCount == 0) || (rowCounter.GCFailCount == 1))
            {
                ReviseExcelValue(SheetEnum.yieldSheet, GCindex, 2, int.Parse(excelDataModel_get.YieldSheet_GC_Input));
                ReviseExcelValue(SheetEnum.yieldSheet, GCindex, 3, int.Parse(excelDataModel_get.YieldSheet_GC_Pass));
                ReviseExcelValue(SheetEnum.yieldSheet, GCindex, 4, int.Parse(excelDataModel_get.YieldSheet_GC_Fail));
                sheet.GetRow(GCindex).GetCell(5).SetCellFormula(String.Format("D{0:G}/C{1:G}", GCindex + 1, GCindex + 1));
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

            if ((rowCounter.FFFailCount == 0) || (rowCounter.FFFailCount == 1))
            {
                ReviseExcelValue(SheetEnum.yieldSheet, FFindex, 2, int.Parse(excelDataModel_get.YieldSheet_FF_Input));
                ReviseExcelValue(SheetEnum.yieldSheet, FFindex, 3, int.Parse(excelDataModel_get.YieldSheet_FF_Pass));
                ReviseExcelValue(SheetEnum.yieldSheet, FFindex, 4, int.Parse(excelDataModel_get.YieldSheet_FF_Fail));
                sheet.GetRow(FFindex).GetCell(5).SetCellFormula(String.Format("D{0:G}/C{1:G}", FFindex + 1, FFindex + 1));
            }
            else
            {
                ReviseExcelValue(SheetEnum.yieldSheet, FFindex, 2, int.Parse(excelDataModel_get.YieldSheet_FF_Input));
                ReviseExcelValue(SheetEnum.yieldSheet, FFindex, 3, int.Parse(excelDataModel_get.YieldSheet_FF_Pass));
                ReviseExcelValue(SheetEnum.yieldSheet, FFindex, 4, int.Parse(excelDataModel_get.YieldSheet_FF_Fail));
                sheet.GetRow(FFindex).GetCell(5).SetCellFormula(String.Format("D{0:G}/C{1:G}", FFindex + 1, FFindex + 1));

                sheet.ShiftRows(GTindex, sheet.LastRowNum, rowCounter.FFFailCount - 1, true, false);
                GTindex += rowCounter.FFFailCount - 1;
                GT2index += rowCounter.FFFailCount - 1;
                CellRangeAddress region;

                for (int i = 1; i <= rowCounter.FFFailCount - 1; i++)
                {
                    sheet.CreateRow(FFindex + i);
                }

                for (int j = 0; j < 10; j++)
                {
                    for (int i = 1; i <= rowCounter.FFFailCount - 1; i++)
                    {
                        SetCellBorderStyle(SheetEnum.yieldSheet, FFindex + i, j + 1);
                    }
                }

                for (int i = 0; i < 5; i++)
                {
                    region = new CellRangeAddress(FFindex, FFindex + rowCounter.FFFailCount - 1, i + 1, i + 1);
                    sheet.AddMergedRegion(region);
                }
            }

            if ((rowCounter.GTFailCount == 0) || (rowCounter.GTFailCount == 1))
            {
                ReviseExcelValue(SheetEnum.yieldSheet, GTindex, 2, int.Parse(excelDataModel_get.YieldSheet_GT_Input));
                ReviseExcelValue(SheetEnum.yieldSheet, GTindex, 3, int.Parse(excelDataModel_get.YieldSheet_GT_Pass));
                ReviseExcelValue(SheetEnum.yieldSheet, GTindex, 4, int.Parse(excelDataModel_get.YieldSheet_GT_Fail));
                sheet.GetRow(GTindex).GetCell(5).SetCellFormula(String.Format("D{0:G}/C{1:G}", GTindex + 1, GTindex + 1));
            }
            else
            {
                ReviseExcelValue(SheetEnum.yieldSheet, GTindex, 2, int.Parse(excelDataModel_get.YieldSheet_GT_Input));
                ReviseExcelValue(SheetEnum.yieldSheet, GTindex, 3, int.Parse(excelDataModel_get.YieldSheet_GT_Pass));
                ReviseExcelValue(SheetEnum.yieldSheet, GTindex, 4, int.Parse(excelDataModel_get.YieldSheet_GT_Fail));
                sheet.GetRow(GTindex).GetCell(5).SetCellFormula(String.Format("D{0:G}/C{1:G}", GTindex + 1, GTindex + 1));

                sheet.ShiftRows(GT2index, sheet.LastRowNum, rowCounter.GTFailCount - 1, true, false);
                GT2index += rowCounter.GTFailCount - 1;
                CellRangeAddress region;

                for (int i = 1; i <= rowCounter.GTFailCount - 1; i++)
                {
                    sheet.CreateRow(GTindex + i);
                }

                for (int j = 0; j < 10; j++)
                {
                    for (int i = 1; i <= rowCounter.GTFailCount - 1; i++)
                    {
                        SetCellBorderStyle(SheetEnum.yieldSheet, GTindex + i, j + 1);
                    }
                }

                for (int i = 0; i < 5; i++)
                {
                    region = new CellRangeAddress(GTindex, GTindex + rowCounter.GTFailCount - 1, i + 1, i + 1);
                    sheet.AddMergedRegion(region);
                }
            }

            if ((rowCounter.GT2FailCount == 0) || (rowCounter.GT2FailCount == 1))
            {
                ReviseExcelValue(SheetEnum.yieldSheet, GT2index, 2, int.Parse(excelDataModel_get.YieldSheet_GT2_Input));
                ReviseExcelValue(SheetEnum.yieldSheet, GT2index, 3, int.Parse(excelDataModel_get.YieldSheet_GT2_Pass));
                ReviseExcelValue(SheetEnum.yieldSheet, GT2index, 4, int.Parse(excelDataModel_get.YieldSheet_GT2_Fail));
                sheet.GetRow(GT2index).GetCell(5).SetCellFormula(String.Format("D{0:G}/C{1:G}", GT2index + 1, GT2index + 1));
            }
            else
            {
                ReviseExcelValue(SheetEnum.yieldSheet, GT2index, 2, int.Parse(excelDataModel_get.YieldSheet_GT2_Input));
                ReviseExcelValue(SheetEnum.yieldSheet, GT2index, 3, int.Parse(excelDataModel_get.YieldSheet_GT2_Pass));
                ReviseExcelValue(SheetEnum.yieldSheet, GT2index, 4, int.Parse(excelDataModel_get.YieldSheet_GT2_Fail));
                sheet.GetRow(GT2index).GetCell(5).SetCellFormula(String.Format("D{0:G}/C{1:G}", GT2index + 1, GT2index + 1));

                //sheet.ShiftRows(GT2index, sheet.LastRowNum, rowCounter.GTFailCount - 1, true, false);
                //GT2index += rowCounter.GTFailCount - 1;
                CellRangeAddress region;

                for (int i = 1; i <= rowCounter.GT2FailCount - 1; i++)
                {
                    sheet.CreateRow(GT2index + i);
                }

                for (int j = 0; j < 10; j++)
                {
                    for (int i = 1; i <= rowCounter.GT2FailCount - 1; i++)
                    {
                        SetCellBorderStyle(SheetEnum.yieldSheet, GT2index + i, j + 1);
                    }
                }

                for (int i = 0; i < 5; i++)
                {
                    region = new CellRangeAddress(GT2index, GT2index + rowCounter.GT2FailCount - 1, i + 1, i + 1);
                    sheet.AddMergedRegion(region);
                }
            }

            using (FileStream fileStream = File.Open(_filePath, FileMode.Create, FileAccess.Write))
            {
                sheet.ForceFormulaRecalculation = true;
                wb.Write(fileStream);
                GCindex = 3;
                FFindex = 4;
                GTindex = 5;
                GT2index = 6;
            }
        }

        

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
    }
}
