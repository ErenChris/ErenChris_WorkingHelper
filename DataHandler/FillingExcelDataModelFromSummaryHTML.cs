using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WorkingHelper.Models;

namespace WorkingHelper.Handler
{
    class FillingExcelDataModelFromSummaryHTML : IFillingExcelDataModelFromSummaryHTML
    {
        HtmlDocument HtmlDocument_Summary = new HtmlDocument();
        HtmlNode node;
        ExcelDataFromSummaryHTMLModel excelDataModel = new ExcelDataFromSummaryHTMLModel();
        ETextReader TR_Summary = new ETextReader(@"D:\Desktop\ALL.txt");

        string[] xPath = new string[] {
            "//*[@id=\"analysis-drop-container\"]/div/div[3]/div/div[2]/div/table/tbody/tr[1]/td[1]",
            "//*[@id=\"analysis-drop-container\"]/div/div[3]/div/div[2]/div/table/tbody/tr[2]/td[1]",
            "//*[@id=\"analysis-drop-container\"]/div/div[3]/div/div[2]/div/table/tbody/tr[3]/td[1]",
            "//*[@id=\"analysis-drop-container\"]/div/div[3]/div/div[2]/div/table/tbody/tr[4]/td[1]"
        };
        string InputCountXPath = "/div/div[2]";
        string FailCountXPath = "/div/div[3]";
        string PassCountXPath = "/div/div[4]";
        string RetestCountXPath = "/div/div[2]";

        public ExcelDataFromSummaryHTMLModel StartCheckStation()
        {
            string result = TR_Summary.GatTextFile();
            HtmlDocument_Summary.LoadHtml(result);

            for (int i = 0; i < 4; i++)
            {
                node = HtmlDocument_Summary.DocumentNode.SelectSingleNode(xPath[i]);
                if (node.Attributes["value"].Value == "GRPC|GRAPE-CAL") FillingGCData(xPath[i]);
                else if (node.Attributes["value"].Value == "FLYT|FIREFLY-TEST") FillingFFData(xPath[i]);
                else if (node.Attributes["value"].Value == "GPST|GRAPE-SELFTEST") FillingGT2Data(xPath[i]);
                else if (node.Attributes["value"].Value == "TGRP|GRAPE-TEST") FillingGTData(xPath[i]);
                else Console.WriteLine("No station match!");
            }

            return excelDataModel;
        }

        //private void FillingAction(string FinalXPath, string Modle)
        //{
        //    node = HtmlDocumentContainer.DocumentNode.SelectSingleNode(FinalXPath);
            
        //}

        private void FillingGCData(string DataPath)
        {
            string HTMLOfAllYieldPath = TextAndXpathHandler.ChangeToChildXPath(DataPath, 2);
            string HTMLOfAllRetestPath = TextAndXpathHandler.ChangeToChildXPath(DataPath, 3);
            node = HtmlDocument_Summary.DocumentNode.SelectSingleNode(HTMLOfAllYieldPath + InputCountXPath);
            excelDataModel.YieldSheet_GC_Input = node.Attributes["data-value"].Value;
            node = HtmlDocument_Summary.DocumentNode.SelectSingleNode(HTMLOfAllYieldPath + FailCountXPath);
            excelDataModel.YieldSheet_GC_Fail = node.Attributes["data-value"].Value;
            node = HtmlDocument_Summary.DocumentNode.SelectSingleNode(HTMLOfAllYieldPath + PassCountXPath);
            excelDataModel.YieldSheet_GC_Pass = node.Attributes["data-value"].Value;

            node = HtmlDocument_Summary.DocumentNode.SelectSingleNode(HTMLOfAllRetestPath + RetestCountXPath);
            excelDataModel.RetestSheet_GC_RetestCount = node.Attributes["data-value"].Value;
        }

        private void FillingFFData(string DataPath)
        {
            string HTMLOfAllYieldPath = TextAndXpathHandler.ChangeToChildXPath(DataPath, 2);
            string HTMLOfAllRetestPath = TextAndXpathHandler.ChangeToChildXPath(DataPath, 3);
            node = HtmlDocument_Summary.DocumentNode.SelectSingleNode(HTMLOfAllYieldPath + InputCountXPath);
            excelDataModel.YieldSheet_FF_Input = node.Attributes["data-value"].Value;
            node = HtmlDocument_Summary.DocumentNode.SelectSingleNode(HTMLOfAllYieldPath + FailCountXPath);
            excelDataModel.YieldSheet_FF_Fail = node.Attributes["data-value"].Value;
            node = HtmlDocument_Summary.DocumentNode.SelectSingleNode(HTMLOfAllYieldPath + PassCountXPath);
            excelDataModel.YieldSheet_FF_Pass = node.Attributes["data-value"].Value;

            node = HtmlDocument_Summary.DocumentNode.SelectSingleNode(HTMLOfAllRetestPath + RetestCountXPath);
            excelDataModel.RetestSheet_FF_RetestCount = node.Attributes["data-value"].Value;
        }

        private void FillingGTData(string DataPath)
        {
            string HTMLOfAllYieldPath = TextAndXpathHandler.ChangeToChildXPath(DataPath, 2);
            string HTMLOfAllRetestPath = TextAndXpathHandler.ChangeToChildXPath(DataPath, 3);
            node = HtmlDocument_Summary.DocumentNode.SelectSingleNode(HTMLOfAllYieldPath + InputCountXPath);
            excelDataModel.YieldSheet_GT_Input = node.Attributes["data-value"].Value;
            node = HtmlDocument_Summary.DocumentNode.SelectSingleNode(HTMLOfAllYieldPath + FailCountXPath);
            excelDataModel.YieldSheet_GT_Fail = node.Attributes["data-value"].Value;
            node = HtmlDocument_Summary.DocumentNode.SelectSingleNode(HTMLOfAllYieldPath + PassCountXPath);
            excelDataModel.YieldSheet_GT_Pass = node.Attributes["data-value"].Value;

            node = HtmlDocument_Summary.DocumentNode.SelectSingleNode(HTMLOfAllRetestPath + RetestCountXPath);
            excelDataModel.RetestSheet_GT_RetestCount = node.Attributes["data-value"].Value;
        }

        private void FillingGT2Data(string DataPath)
        {
            string HTMLOfAllYieldPath = TextAndXpathHandler.ChangeToChildXPath(DataPath, 2);
            string HTMLOfAllRetestPath = TextAndXpathHandler.ChangeToChildXPath(DataPath, 3);
            node = HtmlDocument_Summary.DocumentNode.SelectSingleNode(HTMLOfAllYieldPath + InputCountXPath);
            excelDataModel.YieldSheet_GT2_Input = node.Attributes["data-value"].Value;
            node = HtmlDocument_Summary.DocumentNode.SelectSingleNode(HTMLOfAllYieldPath + FailCountXPath);
            excelDataModel.YieldSheet_GT2_Fail = node.Attributes["data-value"].Value;
            node = HtmlDocument_Summary.DocumentNode.SelectSingleNode(HTMLOfAllYieldPath + PassCountXPath);
            excelDataModel.YieldSheet_GT2_Pass = node.Attributes["data-value"].Value;

            node = HtmlDocument_Summary.DocumentNode.SelectSingleNode(HTMLOfAllRetestPath + RetestCountXPath);
            excelDataModel.RetestSheet_GT2_RetestCount = node.Attributes["data-value"].Value;
        }
    }
}
