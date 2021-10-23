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
using WorkingHelper.Tools;

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

        public void ReviseExcelValue(SheetEnum sheetEnum, int rowindex, int colindex, string context)
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

        public List<RetestUnitModel> DeleteNullFailItem(List<RetestUnitModel> retestUnitModel)
        {
            for (int i = 0; i < retestUnitModel.Count; i++)
            {
                if (retestUnitModel[i].RetestItem.Trim() == "-")
                {
                    retestUnitModel.RemoveAt(i);
                }
            }

            return retestUnitModel;
        }

        public void RetestSheetFilling(RowCounter rowCounter, ExcelDataFromSummaryHTMLModel excelDataModel_get, params List<RetestUnitModel>[] retestUnitModels)
        {
            ISheet sheet = wb.GetSheetAt((int)SheetEnum.retestSheet);

            for (int i = 0; i < retestUnitModels.Length; i++)
            {
                retestUnitModels[i] = DeleteNullFailItem(retestUnitModels[i]);
            }

            IEnumerable<IGrouping<string, RetestUnitModel>> GCRetestUnitsGroupQuery = GeneralTools.GetRetestUnitsGroupQuery(retestUnitModels[0]);
            IEnumerable<IGrouping<string, RetestUnitModel>> FFRetestUnitsGroupQuery = GeneralTools.GetRetestUnitsGroupQuery(retestUnitModels[1]);
            IEnumerable<IGrouping<string, RetestUnitModel>> GTRetestUnitsGroupQuery = GeneralTools.GetRetestUnitsGroupQuery(retestUnitModels[2]);
            IEnumerable<IGrouping<string, RetestUnitModel>> GT2RetestUnitsGroupQuery = GeneralTools.GetRetestUnitsGroupQuery(retestUnitModels[3]);
            int GCRetestCount = GCRetestUnitsGroupQuery.Count();
            int FFRetestCount = FFRetestUnitsGroupQuery.Count();
            int GTRetestCount = GTRetestUnitsGroupQuery.Count();
            int GT2RetestCount = GT2RetestUnitsGroupQuery.Count();
            Console.WriteLine(GCRetestCount.ToString());
            Console.WriteLine(FFRetestCount.ToString());
            Console.WriteLine(GTRetestCount.ToString());
            Console.WriteLine(GT2RetestCount.ToString());

            #region
            if ((retestUnitModels[0].Count == 0) || (retestUnitModels[0].Count == 1))
            {
                ReviseExcelValue(SheetEnum.retestSheet, GCindex, 2, int.Parse(excelDataModel_get.YieldSheet_GC_Input));
                ReviseExcelValue(SheetEnum.retestSheet, GCindex, 3, retestUnitModels[0].Count);
                sheet.GetRow(GCindex).GetCell(4).SetCellFormula(String.Format("D{0:G}/C{1:G}", GCindex + 1, GCindex + 1));
                ReviseExcelValue(SheetEnum.retestSheet, GCindex, 5, retestUnitModels[0].Count);
                sheet.GetRow(GCindex).GetCell(6).SetCellFormula(String.Format("F{0:G}/C{1:G}", GCindex + 1, GCindex + 1));
                ReviseExcelValue(SheetEnum.retestSheet, GCindex, 7, retestUnitModels[0].First<RetestUnitModel>().RetestItem);
                ReviseExcelValue(SheetEnum.retestSheet, GCindex, 8, retestUnitModels[0].First<RetestUnitModel>().RetestStationID);
                ReviseExcelValue(SheetEnum.retestSheet, GCindex, 9, retestUnitModels[0].First<RetestUnitModel>().RetestUnitSN);
                ReviseExcelValue(SheetEnum.retestSheet, GCindex, 10, retestUnitModels[0].First<RetestUnitModel>().UnitConfig);
            }
            else
            {
                ReviseExcelValue(SheetEnum.retestSheet, GCindex, 2, int.Parse(excelDataModel_get.YieldSheet_GC_Input));
                ReviseExcelValue(SheetEnum.retestSheet, GCindex, 3, retestUnitModels[0].Count);
                sheet.GetRow(GCindex).GetCell(4).SetCellFormula(String.Format("D{0:G}/C{1:G}", GCindex + 1, GCindex + 1));

                //string a = String.Format("D{0:G}/C{1:G}", GCindex, GCindex);

                sheet.ShiftRows(FFindex, sheet.LastRowNum, retestUnitModels[0].Count - 1, true, false);
                FFindex += retestUnitModels[0].Count - 1;
                GTindex += retestUnitModels[0].Count - 1;
                GT2index += retestUnitModels[0].Count - 1;
                CellRangeAddress region;

                for (int i = 1; i <= retestUnitModels[0].Count - 1; i++)
                {
                    sheet.CreateRow(GCindex + i);
                }

                for (int j = 0; j < 13; j++)
                {
                    for (int i = 1; i <= retestUnitModels[0].Count - 1; i++)
                    {
                        SetCellBorderStyle(SheetEnum.retestSheet, GCindex + i, j + 1);
                    }
                }

                for (int i = 0; i < 4; i++)
                {
                    region = new CellRangeAddress(GCindex, GCindex + retestUnitModels[0].Count - 1, i + 1, i + 1);
                    sheet.AddMergedRegion(region);
                }
            }
            #endregion

            if ((retestUnitModels[1].Count == 0) || (retestUnitModels[1].Count == 1))
            {
                ReviseExcelValue(SheetEnum.retestSheet, FFindex, 2, int.Parse(excelDataModel_get.YieldSheet_FF_Input));
                ReviseExcelValue(SheetEnum.retestSheet, FFindex, 3, retestUnitModels[1].Count);
                sheet.GetRow(FFindex).GetCell(4).SetCellFormula(String.Format("D{0:G}/C{1:G}", FFindex + 1, FFindex + 1));
                ReviseExcelValue(SheetEnum.retestSheet, FFindex, 5, retestUnitModels[1].Count);
                sheet.GetRow(FFindex).GetCell(6).SetCellFormula(String.Format("F{0:G}/C{1:G}", FFindex + 1, FFindex + 1));
                ReviseExcelValue(SheetEnum.retestSheet, FFindex, 7, retestUnitModels[1].First<RetestUnitModel>().RetestItem);
                ReviseExcelValue(SheetEnum.retestSheet, FFindex, 8, retestUnitModels[1].First<RetestUnitModel>().RetestStationID);
                ReviseExcelValue(SheetEnum.retestSheet, FFindex, 9, retestUnitModels[1].First<RetestUnitModel>().RetestUnitSN);
                ReviseExcelValue(SheetEnum.retestSheet, FFindex, 10, retestUnitModels[1].First<RetestUnitModel>().UnitConfig);
            }
            else
            {
                ReviseExcelValue(SheetEnum.retestSheet, FFindex, 2, int.Parse(excelDataModel_get.YieldSheet_FF_Input));
                ReviseExcelValue(SheetEnum.retestSheet, FFindex, 3, retestUnitModels[1].Count);
                sheet.GetRow(FFindex).GetCell(4).SetCellFormula(String.Format("D{0:G}/C{1:G}", FFindex + 1, FFindex + 1));

                //string a = String.Format("D{0:G}/C{1:G}", GCindex, GCindex);

                sheet.ShiftRows(GTindex, sheet.LastRowNum, retestUnitModels[1].Count - 1, true, false);
                GTindex += retestUnitModels[1].Count - 1;
                GT2index += retestUnitModels[1].Count - 1;
                CellRangeAddress region;

                for (int i = 1; i <= retestUnitModels[1].Count - 1; i++)
                {
                    sheet.CreateRow(FFindex + i);
                }

                for (int j = 0; j < 13; j++)
                {
                    for (int i = 1; i <= retestUnitModels[1].Count - 1; i++)
                    {
                        SetCellBorderStyle(SheetEnum.retestSheet, FFindex + i, j + 1);
                    }
                }

                for (int i = 0; i < 4; i++)
                {
                    region = new CellRangeAddress(FFindex, FFindex + retestUnitModels[1].Count - 1, i + 1, i + 1);
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
    }
}
