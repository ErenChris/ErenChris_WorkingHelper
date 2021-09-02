using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorkingHelper.Models
{
    class ExcelDataFromSummaryHTMLModel
    {
        public string YieldSheet_GC_Input { get; set; }
        public string YieldSheet_GC_Pass { get; set; }
        public string YieldSheet_GC_Fail { get; set; }

        public string YieldSheet_FF_Input { get; set; }
        public string YieldSheet_FF_Pass { get; set; }
        public string YieldSheet_FF_Fail { get; set; }

        public string YieldSheet_GT_Input { get; set; }
        public string YieldSheet_GT_Pass { get; set; }
        public string YieldSheet_GT_Fail { get; set; }

        public string YieldSheet_GT2_Input { get; set; }
        public string YieldSheet_GT2_Pass { get; set; }
        public string YieldSheet_GT2_Fail { get; set; }

        public string RetestSheet_GC_RetestCount { get; set; }

        public string RetestSheet_FF_RetestCount { get; set; }

        public string RetestSheet_GT_RetestCount { get; set; }

        public string RetestSheet_GT2_RetestCount { get; set; }
    }
}
