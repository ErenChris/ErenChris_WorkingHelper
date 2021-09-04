using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WorkingHelper.Handler;

namespace WorkingHelper.Models
{
    class RowCounter
    {
        public int GCFailCount { get; set; }
        public int FFFailCount { get; set; }
        public int GTFailCount { get; set; }
        public int GT2FailCount { get; set; }

        public int GCRetestCount { get; set; }
        public int FFRetestCount { get; set; }
        public int GTRetestCount { get; set; }
        public int GT2RetestCount { get; set; }

        public RowCounter(List<RetestUnitModel> GCRetest, List<RetestUnitModel> FFRetest, List<RetestUnitModel> GTRetest, List<RetestUnitModel> GT2Retest, ExcelDataFromSummaryHTMLModel excelDataModel)
        {
            GCRetestCount = GCRetest.Count;
            FFRetestCount = FFRetest.Count;
            GTRetestCount = GTRetest.Count;
            GT2RetestCount = GT2Retest.Count;

            GCFailCount = int.Parse(excelDataModel.YieldSheet_GC_Fail);
            FFFailCount = int.Parse(excelDataModel.YieldSheet_FF_Fail);
            GTFailCount = int.Parse(excelDataModel.YieldSheet_GT_Fail);
            GT2FailCount = int.Parse(excelDataModel.YieldSheet_GT2_Fail);
        }
    }
}
