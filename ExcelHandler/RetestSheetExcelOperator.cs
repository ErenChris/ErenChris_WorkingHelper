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
            var resourseRowStyle = sheet.GetRow(GCindex).RowStyle;
            var resourseRowForSetStyle = sheet.GetRow(GCindex);

            for (int i = 0; i < retestUnitModels.Length; i++)
            {
                retestUnitModels[i] = DeleteNullFailItem(retestUnitModels[i]);
            }

            //var query = retestUnitModels[3].OrderBy(p => new { p.RetestItem, p.RetestStationID });

            IEnumerable<IGrouping<string, RetestUnitModel>> GCRetestUnitsGroupQuery = GeneralTools.GetRetestUnitsGroupQuery(retestUnitModels[0]);
            IEnumerable<IGrouping<string, RetestUnitModel>> FFRetestUnitsGroupQuery = GeneralTools.GetRetestUnitsGroupQuery(retestUnitModels[1]);
            IEnumerable<IGrouping<string, RetestUnitModel>> GTRetestUnitsGroupQuery = GeneralTools.GetRetestUnitsGroupQuery(retestUnitModels[2]);
            IEnumerable<IGrouping<string, RetestUnitModel>> GT2RetestUnitsGroupQuery = GeneralTools.GetRetestUnitsGroupQuery(retestUnitModels[3]);
            
            int GCRetestGroupCount = GCRetestUnitsGroupQuery.Count();
            int FFRetestGroupCount = FFRetestUnitsGroupQuery.Count();
            int GTRetestGroupCount = GTRetestUnitsGroupQuery.Count();
            int GT2RetestGroupCount = GT2RetestUnitsGroupQuery.Count();

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

                sheet.ShiftRows(FFindex, sheet.LastRowNum, GCRetestGroupCount - 1, true, false);
                FFindex += GCRetestGroupCount - 1;
                GTindex += GCRetestGroupCount - 1;
                GT2index += GCRetestGroupCount - 1;
                CellRangeAddress region;

                for (int i = 1; i <= GCRetestGroupCount - 1; i++)
                {
                    sheet.CreateRow(GCindex + i);
                    for (int j = 1; j < 11; j++)
                    {
                        IRow row = sheet.GetRow(GCindex + i);
                        row.CreateCell(j);
                    }
                }

                for (int i = 1; i <= GCRetestGroupCount - 1; i++)
                {
                    for (int j = 0; j < 13; j++)
                    {
                        SetCellBorderStyle(SheetEnum.retestSheet, GCindex + i, j + 1);
                    }
                }

                for (int i = 0; i < 4; i++)
                {
                    region = new CellRangeAddress(GCindex, GCindex + GCRetestGroupCount - 1, i + 1, i + 1);
                    sheet.AddMergedRegion(region);
                }

                //for (int i = 0; i < GCRetestGroupCount; i++)
                //{
                //    for (int j = 0; j < GCRetestUnitsGroupQuery.; j++)
                //    {

                //    }
                //}
                int GCTempCount = 0;
                List<RetestUnitModel> GCSingleRetestUnit = new List<RetestUnitModel>();
                foreach (var group in GCRetestUnitsGroupQuery)
                {
                    //分组之后，要填充整行
                    //填充整行时要遍历单个机台
                    //对组内机台进行分组

                    //未创建cell实例，引起第二次循环报错(已解决)
                    ReviseExcelValue(SheetEnum.retestSheet, GCindex + GCTempCount, 5, group.Count());
                    sheet.GetRow(GCindex + GCTempCount).GetCell(6).SetCellFormula(String.Format("F{0:G}/C{1:G}", GCindex + 1 + GCTempCount, GCindex + 1));
                    ReviseExcelValue(SheetEnum.retestSheet, GCindex + GCTempCount, 7, group.Key);

                    if (group.Count() > 1)
                    {
                        string strRetestStation = null;
                        string strRetestSN = null;
                        string strRetestConfig = null;
                        int flag = 0;

                        foreach (var i in group)
                        {
                            GCSingleRetestUnit.Add(i);
                        }
                        IEnumerable<IGrouping<string, RetestUnitModel>> GCRetestUnitsGroupQueryByStation = GeneralTools.GetRetestUnitsGroupQueryByStation(GCSingleRetestUnit);
                        foreach (var groupByStation in GCRetestUnitsGroupQueryByStation)
                        {
                            if (flag == 0)
                            {
                                strRetestStation += groupByStation.First().RetestStationID + String.Format(" x{0:G}", groupByStation.Count());
                                strRetestSN += groupByStation.First().RetestUnitSN;
                                strRetestConfig += groupByStation.First().UnitConfig;
                            }
                            else
                            {
                                strRetestStation = strRetestStation + "\n" + groupByStation.First().RetestStationID;
                                strRetestSN = strRetestSN + "\n" + groupByStation.First().RetestUnitSN;
                                strRetestConfig = strRetestConfig + "\n" + groupByStation.First().UnitConfig;
                            }
                            flag += 1;
                        }
                        ReviseExcelValue(SheetEnum.retestSheet, GCindex + GCTempCount, 8, strRetestStation);
                        ReviseExcelValue(SheetEnum.retestSheet, GCindex + GCTempCount, 9, strRetestSN);
                        ReviseExcelValue(SheetEnum.retestSheet, GCindex + GCTempCount, 10, strRetestConfig);
                    }
                    else
                    {
                        foreach (var i in group)
                        {
                            ReviseExcelValue(SheetEnum.retestSheet, GCindex + GCTempCount, 8, i.RetestStationID);
                            ReviseExcelValue(SheetEnum.retestSheet, GCindex + GCTempCount, 9, i.RetestUnitSN);
                            ReviseExcelValue(SheetEnum.retestSheet, GCindex + GCTempCount, 10, i.UnitConfig);
                        }
                    }
                    GCTempCount += 1;
                    GCSingleRetestUnit.Clear();
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

                sheet.ShiftRows(GTindex, sheet.LastRowNum, FFRetestGroupCount - 1, true, false);
                GTindex += FFRetestGroupCount - 1;
                GT2index += FFRetestGroupCount - 1;
                CellRangeAddress region;

                //for (int i = 1; i <= retestUnitModels[1].Count - 1; i++)
                //{
                //    sheet.CreateRow(FFindex + i);
                //}
                for (int i = 1; i <= FFRetestGroupCount - 1; i++)
                {
                    sheet.CreateRow(FFindex + i);
                    for (int j = 1; j < 11; j++)
                    {
                        IRow row = sheet.GetRow(FFindex + i);
                        row.CreateCell(j);
                    }
                }

                for (int i = 1; i <= FFRetestGroupCount - 1; i++)
                {
                    for (int j = 0; j < 13; j++)
                    {
                        SetCellBorderStyle(SheetEnum.retestSheet, FFindex + i, j + 1);
                    }
                }

                for (int i = 0; i < 4; i++)
                {
                    region = new CellRangeAddress(FFindex, FFindex + FFRetestGroupCount - 1, i + 1, i + 1);
                    sheet.AddMergedRegion(region);
                }

                int FFTempCount = 0;
                List<RetestUnitModel> FFSingleRetestUnit = new List<RetestUnitModel>();
                foreach (var group in FFRetestUnitsGroupQuery)
                {
                    //分组之后，要填充整行
                    //填充整行时要遍历单个机台
                    //对组内机台进行分组

                    //未创建cell实例，引起第二次循环报错(已解决)
                    ReviseExcelValue(SheetEnum.retestSheet, FFindex + FFTempCount, 5, group.Count());
                    sheet.GetRow(FFindex + FFTempCount).GetCell(6).SetCellFormula(String.Format("F{0:G}/C{1:G}", FFindex + 1 + FFTempCount, FFindex + 1));
                    ReviseExcelValue(SheetEnum.retestSheet, FFindex + FFTempCount, 7, group.Key);

                    if (group.Count() > 1)
                    {
                        string strRetestStation = null;
                        string strRetestSN = null;
                        string strRetestConfig = null;
                        int flag = 0;

                        foreach (var i in group)
                        {
                            FFSingleRetestUnit.Add(i);
                        }
                        IEnumerable<IGrouping<string, RetestUnitModel>> FFRetestUnitsGroupQueryByStation = GeneralTools.GetRetestUnitsGroupQueryByStation(FFSingleRetestUnit);
                        foreach (var groupByStation in FFRetestUnitsGroupQueryByStation)
                        {
                            if (flag == 0)
                            {
                                strRetestStation += groupByStation.First().RetestStationID + String.Format(" x{0:G}", groupByStation.Count());
                                strRetestSN += groupByStation.First().RetestUnitSN;
                                strRetestConfig += groupByStation.First().UnitConfig;
                            }
                            else
                            {
                                strRetestStation = strRetestStation + "\n" + groupByStation.First().RetestStationID;
                                strRetestSN = strRetestSN + "\n" + groupByStation.First().RetestUnitSN;
                                strRetestConfig = strRetestConfig + "\n" + groupByStation.First().UnitConfig;
                            }
                            flag += 1;
                        }
                        ReviseExcelValue(SheetEnum.retestSheet, FFindex + FFTempCount, 8, strRetestStation);
                        ReviseExcelValue(SheetEnum.retestSheet, FFindex + FFTempCount, 9, strRetestSN);
                        ReviseExcelValue(SheetEnum.retestSheet, FFindex + FFTempCount, 10, strRetestConfig);
                    }
                    else
                    {
                        foreach (var i in group)
                        {
                            ReviseExcelValue(SheetEnum.retestSheet, FFindex + FFTempCount, 8, i.RetestStationID);
                            ReviseExcelValue(SheetEnum.retestSheet, FFindex + FFTempCount, 9, i.RetestUnitSN);
                            ReviseExcelValue(SheetEnum.retestSheet, FFindex + FFTempCount, 10, i.UnitConfig);
                        }
                    }
                    FFTempCount += 1;
                    FFSingleRetestUnit.Clear();
                }
            }

            if ((retestUnitModels[2].Count == 0) || (retestUnitModels[2].Count == 1))
            {
                ReviseExcelValue(SheetEnum.retestSheet, GTindex, 2, int.Parse(excelDataModel_get.YieldSheet_GT_Input));
                ReviseExcelValue(SheetEnum.retestSheet, GTindex, 3, retestUnitModels[2].Count);
                sheet.GetRow(GTindex).GetCell(4).SetCellFormula(String.Format("D{0:G}/C{1:G}", GTindex + 1, GTindex + 1));
                ReviseExcelValue(SheetEnum.retestSheet, GTindex, 5, retestUnitModels[2].Count);
                sheet.GetRow(GTindex).GetCell(6).SetCellFormula(String.Format("F{0:G}/C{1:G}", GTindex + 1, GTindex + 1));
                ReviseExcelValue(SheetEnum.retestSheet, GTindex, 7, retestUnitModels[2].First<RetestUnitModel>().RetestItem);
                ReviseExcelValue(SheetEnum.retestSheet, GTindex, 8, retestUnitModels[2].First<RetestUnitModel>().RetestStationID);
                ReviseExcelValue(SheetEnum.retestSheet, GTindex, 9, retestUnitModels[2].First<RetestUnitModel>().RetestUnitSN);
                ReviseExcelValue(SheetEnum.retestSheet, GTindex, 10, retestUnitModels[2].First<RetestUnitModel>().UnitConfig);
            }
            else
            {
                ReviseExcelValue(SheetEnum.retestSheet, GTindex, 2, int.Parse(excelDataModel_get.YieldSheet_GT_Input));
                ReviseExcelValue(SheetEnum.retestSheet, GTindex, 3, retestUnitModels[2].Count);
                sheet.GetRow(GTindex).GetCell(4).SetCellFormula(String.Format("D{0:G}/C{1:G}", GTindex + 1, GTindex + 1));

                //string a = String.Format("D{0:G}/C{1:G}", GCindex, GCindex);

                sheet.ShiftRows(GT2index, sheet.LastRowNum, GTRetestGroupCount - 1, true, false);
                GT2index += retestUnitModels[2].Count - 1;
                CellRangeAddress region;

                for (int i = 1; i <= GTRetestGroupCount - 1; i++)
                {
                    sheet.CreateRow(GTindex + i);
                    for (int j = 1; j < 11; j++)
                    {
                        IRow row = sheet.GetRow(GTindex + i);
                        row.CreateCell(j);
                    }
                }

                for (int i = 1; i <= GTRetestGroupCount - 1; i++)
                {
                    for (int j = 0; j < 13; j++)
                    {
                        SetCellBorderStyle(SheetEnum.retestSheet, GTindex + i, j + 1);
                    }
                }

                for (int i = 0; i < 4; i++)
                {
                    region = new CellRangeAddress(GTindex, GTindex + GTRetestGroupCount - 1, i + 1, i + 1);
                    sheet.AddMergedRegion(region);
                }

                //for (int i = 0; i < GCRetestGroupCount; i++)
                //{
                //    for (int j = 0; j < GCRetestUnitsGroupQuery.; j++)
                //    {

                //    }
                //}
                int GTTempCount = 0;
                List<RetestUnitModel> GTSingleRetestUnit = new List<RetestUnitModel>();
                foreach (var group in GTRetestUnitsGroupQuery)
                {
                    //分组之后，要填充整行
                    //填充整行时要遍历单个机台
                    //对组内机台进行分组

                    //未创建cell实例，引起第二次循环报错(已解决)
                    ReviseExcelValue(SheetEnum.retestSheet, GTindex + GTTempCount, 5, group.Count());
                    sheet.GetRow(GTindex + GTTempCount).GetCell(6).SetCellFormula(String.Format("F{0:G}/C{1:G}", GTindex + 1 + GTTempCount, GTindex + 1));
                    ReviseExcelValue(SheetEnum.retestSheet, GTindex + GTTempCount, 7, group.Key);

                    if (group.Count() > 1)
                    {
                        string strRetestStation = null;
                        string strRetestSN = null;
                        string strRetestConfig = null;
                        int flag = 0;

                        foreach (var i in group)
                        {
                            GTSingleRetestUnit.Add(i);
                        }
                        IEnumerable<IGrouping<string, RetestUnitModel>> GTRetestUnitsGroupQueryByStation = GeneralTools.GetRetestUnitsGroupQueryByStation(GTSingleRetestUnit);
                        foreach (var groupByStation in GTRetestUnitsGroupQueryByStation)
                        {
                            if (flag == 0)
                            {
                                strRetestStation += groupByStation.First().RetestStationID + String.Format(" x{0:G}", groupByStation.Count());
                                strRetestSN += groupByStation.First().RetestUnitSN;
                                strRetestConfig += groupByStation.First().UnitConfig;
                            }
                            else
                            {
                                strRetestStation = strRetestStation + "\n" + groupByStation.First().RetestStationID;
                                strRetestSN = strRetestSN + "\n" + groupByStation.First().RetestUnitSN;
                                strRetestConfig = strRetestConfig + "\n" + groupByStation.First().UnitConfig;
                            }
                            flag += 1;
                        }
                        ReviseExcelValue(SheetEnum.retestSheet, GTindex + GTTempCount, 8, strRetestStation);
                        ReviseExcelValue(SheetEnum.retestSheet, GTindex + GTTempCount, 9, strRetestSN);
                        ReviseExcelValue(SheetEnum.retestSheet, GTindex + GTTempCount, 10, strRetestConfig);
                    }
                    else
                    {
                        foreach (var i in group)
                        {
                            ReviseExcelValue(SheetEnum.retestSheet, GTindex + GTTempCount, 8, i.RetestStationID);
                            ReviseExcelValue(SheetEnum.retestSheet, GTindex + GTTempCount, 9, i.RetestUnitSN);
                            ReviseExcelValue(SheetEnum.retestSheet, GTindex + GTTempCount, 10, i.UnitConfig);
                        }
                    }
                    GTTempCount += 1;
                    GTSingleRetestUnit.Clear();
                }
            }

            #region
            if ((retestUnitModels[3].Count == 0) || (retestUnitModels[3].Count == 1))
            {
                ReviseExcelValue(SheetEnum.retestSheet, GT2index, 2, int.Parse(excelDataModel_get.YieldSheet_GT2_Input));
                ReviseExcelValue(SheetEnum.retestSheet, GT2index, 3, retestUnitModels[3].Count);
                sheet.GetRow(GT2index).GetCell(4).SetCellFormula(String.Format("D{0:G}/C{1:G}", GT2index + 1, GT2index + 1));
                ReviseExcelValue(SheetEnum.retestSheet, GT2index, 5, retestUnitModels[3].Count);
                sheet.GetRow(GT2index).GetCell(6).SetCellFormula(String.Format("F{0:G}/C{1:G}", GT2index + 1, GT2index + 1));
                ReviseExcelValue(SheetEnum.retestSheet, GT2index, 7, retestUnitModels[3].First<RetestUnitModel>().RetestItem);
                ReviseExcelValue(SheetEnum.retestSheet, GT2index, 8, retestUnitModels[3].First<RetestUnitModel>().RetestStationID);
                ReviseExcelValue(SheetEnum.retestSheet, GT2index, 9, retestUnitModels[3].First<RetestUnitModel>().RetestUnitSN);
                ReviseExcelValue(SheetEnum.retestSheet, GT2index, 10, retestUnitModels[3].First<RetestUnitModel>().UnitConfig);
            }
            else
            {
                ReviseExcelValue(SheetEnum.retestSheet, GT2index, 2, int.Parse(excelDataModel_get.YieldSheet_GT2_Input));
                ReviseExcelValue(SheetEnum.retestSheet, GT2index, 3, retestUnitModels[3].Count);
                sheet.GetRow(GT2index).GetCell(4).SetCellFormula(String.Format("D{0:G}/C{1:G}", GT2index + 1, GT2index + 1));

                //string a = String.Format("D{0:G}/C{1:G}", GCindex, GCindex);

                //sheet.ShiftRows(GT22index, sheet.LastRowNum, retestUnitModels[3].Count - 1, true, false);
                //GT22index += retestUnitModels[3].Count - 1;
                CellRangeAddress region;

                for (int i = 1; i <= GT2RetestGroupCount - 1; i++)
                {
                    sheet.CreateRow(GT2index + i);

                    for (int j = 1; j < 11; j++)
                    {
                        IRow row = sheet.GetRow(GT2index + i);
                        row.CreateCell(j);
                    }
                }

                for (int i = 1; i <= GT2RetestGroupCount - 1; i++)
                {
                    for (int j = 0; j < 13; j++)
                    {
                        SetCellBorderStyle(SheetEnum.retestSheet, GT2index + i, j + 1);
                    }
                }

                for (int i = 0; i < 4; i++)
                {
                    region = new CellRangeAddress(GT2index, GT2index + GT2RetestGroupCount - 1, i + 1, i + 1);
                    sheet.AddMergedRegion(region);
                }

                //for (int i = 0; i < GCRetestGroupCount; i++)
                //{
                //    for (int j = 0; j < GCRetestUnitsGroupQuery.; j++)
                //    {

                //    }
                //}
                int GT2TempCount = 0;
                List<RetestUnitModel> GT2SingleRetestUnit = new List<RetestUnitModel>();
                foreach (var group in GT2RetestUnitsGroupQuery)
                {
                    //分组之后，要填充整行
                    //填充整行时要遍历单个机台
                    //对组内机台进行分组

                    //未创建cell实例，引起第二次循环报错(已解决)
                    ReviseExcelValue(SheetEnum.retestSheet, GT2index + GT2TempCount, 5, group.Count());
                    sheet.GetRow(GT2index + GT2TempCount).GetCell(6).SetCellFormula(String.Format("F{0:G}/C{1:G}", GT2index + 1 + GT2TempCount, GT2index + 1));
                    ReviseExcelValue(SheetEnum.retestSheet, GT2index + GT2TempCount, 7, group.Key);

                    if (group.Count() > 1)
                    {
                        string strRetestStation = null;
                        string strRetestSN = null;
                        string strRetestConfig = null;
                        int flag = 0;

                        foreach (var i in group)
                        {
                            GT2SingleRetestUnit.Add(i);
                        }
                        IEnumerable<IGrouping<string, RetestUnitModel>> GT2RetestUnitsGroupQueryByStation = GeneralTools.GetRetestUnitsGroupQueryByStation(GT2SingleRetestUnit);
                        foreach (var groupByStation in GT2RetestUnitsGroupQueryByStation)
                        {
                            if (flag == 0)
                            {
                                strRetestStation += groupByStation.First().RetestStationID + String.Format(" x{0:G}", groupByStation.Count());
                                strRetestSN += groupByStation.First().RetestUnitSN;
                                strRetestConfig += groupByStation.First().UnitConfig;
                            }
                            else
                            {
                                strRetestStation = strRetestStation + "\n" + groupByStation.First().RetestStationID;
                                strRetestSN = strRetestSN + "\n" + groupByStation.First().RetestUnitSN;
                                strRetestConfig = strRetestConfig + "\n" + groupByStation.First().UnitConfig;
                            }
                            flag += 1;
                        }
                        ReviseExcelValue(SheetEnum.retestSheet, GT2index + GT2TempCount, 8, strRetestStation);
                        ReviseExcelValue(SheetEnum.retestSheet, GT2index + GT2TempCount, 9, strRetestSN);
                        ReviseExcelValue(SheetEnum.retestSheet, GT2index + GT2TempCount, 10, strRetestConfig);
                    }
                    else
                    {
                        foreach (var i in group)
                        {
                            ReviseExcelValue(SheetEnum.retestSheet, GT2index + GT2TempCount, 8, i.RetestStationID);
                            ReviseExcelValue(SheetEnum.retestSheet, GT2index + GT2TempCount, 9, i.RetestUnitSN);
                            ReviseExcelValue(SheetEnum.retestSheet, GT2index + GT2TempCount, 10, i.UnitConfig);
                        }
                    }
                    GT2TempCount += 1;
                    GT2SingleRetestUnit.Clear();
                }
            }
            #endregion

            //ICellStyle cellStylePercentage = wb.CreateCellStyle();
            //cellStylePercentage.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00%");

            using (FileStream fileStream = File.Open(_filePath, FileMode.Create, FileAccess.Write))
            {
                for (int i = 1; i <= (sheet.LastRowNum - GCindex); i++)
                {
                    IRow row = sheet.GetRow(GCindex + i);
                    row.RowStyle = resourseRowStyle;
                    for (int j = 1; j <= 10; j++)
                    {
                        row.GetCell(j).CellStyle = resourseRowForSetStyle.GetCell(j).CellStyle;
                    }
                }

                for (int i = 1; i <= 10; i++)
                {
                    sheet.AutoSizeColumn(i);
                }
                //for (int colNum = 1; colNum <= 10; colNum++)
                //{
                //    int colWidth = sheet.GetColumnWidth(colNum) / 256;
                //    for (int rowNum = 1; rowNum <= sheet.LastRowNum; rowNum++)
                //    {
                //        IRow currentRow = sheet.GetRow(rowNum);
                //        if (currentRow.GetCell(colNum) != null)
                //        {
                //            ICell currentCell = currentRow.GetCell(colNum);
                //            int length = Encoding.Default.GetBytes(currentCell.ToString()).Length;
                //            if (colWidth < length)
                //            {
                //                colWidth = length;
                //            }
                //        }
                //    }
                //    sheet.SetColumnWidth(colNum, colWidth * 256);
                //}

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