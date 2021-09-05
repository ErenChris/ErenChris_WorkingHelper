using NPOI;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.Util;
using System.IO;
using System;
using WorkingHelper.Handler;
using HtmlAgilityPack;
using WorkingHelper.Models;
using System.Collections.Generic;
using WorkingHelper.ExcelHandler;

namespace WorkingHelper
{
    class Program
    {
        static void Main(string[] args)
        {
            string dir = AppDomain.CurrentDomain.BaseDirectory;

            string FileName = dir + @"Resourse\Template.xlsx";

            List<RetestUnitModel> GCRetestUnits = new List<RetestUnitModel>();
            List<RetestUnitModel> FFRetestUnits = new List<RetestUnitModel>();
            List<RetestUnitModel> GTRetestUnits = new List<RetestUnitModel>();
            List<RetestUnitModel> GT2RetestUnits = new List<RetestUnitModel>();
            GCRetestUnits = HTMLToModelOfRetest.GetRetestUnitList(GCRetestUnits, "GC");
            FFRetestUnits = HTMLToModelOfRetest.GetRetestUnitList(FFRetestUnits, "FF");
            GTRetestUnits = HTMLToModelOfRetest.GetRetestUnitList(GTRetestUnits, "GT");
            GT2RetestUnits = HTMLToModelOfRetest.GetRetestUnitList(GT2RetestUnits, "GT2");

            ExcelDataFromSummaryHTMLModel excelDataModel_get;
            IFillingExcelDataModelFromSummaryHTML fillingExcelDataModelFromHTML = new FillingExcelDataModelFromSummaryHTML();
            excelDataModel_get = fillingExcelDataModelFromHTML.StartCheckStation();
            RowCounter rowCounter = new RowCounter(GCRetestUnits, FFRetestUnits, GTRetestUnits, GT2RetestUnits, excelDataModel_get);

            //string Path = @"D:\Desktop\111.xlsx";
            //HtmlDocument HtmlDocumentContainer = new HtmlDocument();
            //HtmlNode node;

            //string xPath = "//*[@id=\"analysis-drop-container\"]/div/div[3]/div/div[2]/div/table/tbody/tr[1]/td[1]";

            //Operate Excel
            //XSSFWorkbook dailyReportWorkBook = new XSSFWorkbook();
            //ISheet sheet = dailyReportWorkBook.CreateSheet("FirstSheet");
            //IRow row = sheet.CreateRow(5);
            //ICell cell = row.CreateCell(3);
            //cell.SetCellValue("测试2");


            ExcelOpreator excelOpreator = new ExcelOpreator(FileName);
            //excelOpreator.ReviseExcelValue(ExcelOpreator.SheetEnum.yieldSheet, 10, 3, "ErenChris");
            int rowsNum = excelOpreator.GetLastRowIndex(ExcelOpreator.SheetEnum.yieldSheet);

            excelOpreator.YieldSheetFilling(rowCounter, excelDataModel_get, GCRetestUnits, FFRetestUnits, GTRetestUnits, GT2RetestUnits);
            //using (FileStream FS = new FileStream(Path, FileMode.Create, FileAccess.Write))
            //{
            //    dailyReportWorkBook.Write(FS);
            //}

            Console.WriteLine("Done!");

            //test txt read function
            //ETextReader TR = new ETextReader(@"D:\Desktop\ALL.txt");
            //string result = TR.GatTextFile();
            //HtmlDocumentContainer.LoadHtml(result);
            //node = HtmlDocumentContainer.GetElementbyId("root");
            //string test = node.FirstChild.Attributes["class"].Value;
            //Console.WriteLine(test);
            //node = HtmlDocumentContainer.DocumentNode.SelectSingleNode(xPath);
            //string test = node.Attributes["value"].Value;
            //Console.WriteLine(test);



            Console.ReadLine();
        }
    }
}
