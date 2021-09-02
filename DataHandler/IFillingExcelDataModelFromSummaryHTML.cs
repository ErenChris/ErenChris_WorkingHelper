using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WorkingHelper.Models;

namespace WorkingHelper.Handler
{
    interface IFillingExcelDataModelFromSummaryHTML
    {
        ExcelDataFromSummaryHTMLModel StartCheckStation();
    }
}
