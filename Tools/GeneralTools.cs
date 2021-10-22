﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WorkingHelper.Models;

namespace WorkingHelper.Tools
{
    public static class GeneralTools
    {
        public static int GetRetestItemsCategoryCount(List<RetestUnitModel> retestUnitModels)
        {
            int count = 0;

            IEnumerable<IGrouping<string, RetestUnitModel>> query = from retestUnitModel in retestUnitModels
                                                                    group retestUnitModel by retestUnitModel.RetestItem;

            foreach(var group in query)
            {
                count += 1;
            }

            return count;
        }

        public static int GetRetestItemsCategoryCount(string[] strArray)
        {
            return 0;
        }
    }
}