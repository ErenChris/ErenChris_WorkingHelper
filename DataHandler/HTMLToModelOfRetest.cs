using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WorkingHelper.Models;

namespace WorkingHelper.Handler
{
    public static class HTMLToModelOfRetest
    {


        public static List<RetestUnitModel> GetRetestUnitList(List<RetestUnitModel> Entity, string modelType)
        {
            HtmlDocument htmlDocument_Retest = new HtmlDocument();
            string readPath = null;

            if (modelType == "GC")
            {
                readPath = @".\Resourse\GCRetest.txt";
            }
            else if (modelType == "FF")
            {
                readPath = @".\Resourse\FFRetest.txt";
            }
            else if (modelType == "GT")
            {
                readPath = @".\Resourse\GTRetest.txt";
            }
            else if (modelType == "GT2")
            {
                readPath = @".\Resourse\GT2Retest.txt";
            }
            else
            {
                Console.WriteLine("GetFFRetestUnitList<T>: No match!");
            }

            ETextReader TR_FFRetest = new ETextReader(readPath);
            string result = TR_FFRetest.GatTextFile();
            htmlDocument_Retest.LoadHtml(result);
            HtmlNode node;

            string xpath = "//*[@id=\"analysis-drop-container\"]/div/div[2]/div/div/table/tbody";
            node = htmlDocument_Retest.DocumentNode.SelectSingleNode(xpath);
            IEnumerable<HtmlNode> nodechildren = node.Elements("tr");
            //int RetestUnitsCount = nodechildren.Count();

            foreach (HtmlNode item in nodechildren)
            {
                RetestUnitModel tempModel = new RetestUnitModel();
                HtmlDocument tempDocument = new HtmlDocument();
                string tempString = item.InnerHtml;
                tempDocument.LoadHtml(tempString);
                HtmlNode tempnode;

                tempnode = tempDocument.DocumentNode.SelectSingleNode("td[1]");
                string testresult = tempnode.Attributes["value"].Value;
                tempModel.RetestUnitSN = testresult;
                tempnode = tempDocument.DocumentNode.SelectSingleNode("td[7]");
                testresult = tempnode.Attributes["value"].Value;
                tempModel.RetestStationID = testresult;
                tempnode = tempDocument.DocumentNode.SelectSingleNode("td[10]");
                testresult = tempnode.Attributes["value"].Value;
                tempModel.RetestItem = testresult;
                tempnode = tempDocument.DocumentNode.SelectSingleNode("td[20]");
                testresult = tempnode.Attributes["value"].Value;
                tempModel.UnitConfig = testresult;


                Entity.Add(tempModel);
            }

            return Entity;
        }
    }
}
