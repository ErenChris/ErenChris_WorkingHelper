using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WorkingHelper.Models;

namespace WorkingHelper.Tools
{
    public static class GeneralTools
    {
        public static IEnumerable<IGrouping<string, RetestUnitModel>> GetRetestUnitsGroupQuery(List<RetestUnitModel> retestUnitModels)
        {
            //int count = 0;

            IEnumerable<IGrouping<string, RetestUnitModel>> query = from retestUnitModel in retestUnitModels
                                                                    group retestUnitModel by retestUnitModel.RetestItem;

            //foreach(var group in query)
            //{
            //    count += 1;
            //}

            return query;
        }

        public static IEnumerable<IGrouping<string, RetestUnitModel>> GetRetestUnitsGroupQueryByStation(List<RetestUnitModel> retestUnitModels)
        {
            //int count = 0;

            IEnumerable<IGrouping<string, RetestUnitModel>> query = from retestUnitModel in retestUnitModels
                                                                    group retestUnitModel by retestUnitModel.RetestStationID;

            //foreach(var group in query)
            //{
            //    count += 1;
            //}

            return query;
        }

        //public static List<RetestUnitModel> GetGroupList(IEnumerable<IGrouping<string, RetestUnitModel>> query)
        //{
        //    List<RetestUnitModel> outPut = new List<RetestUnitModel>();

        //    foreach (var group in query)
        //    {
        //        foreach (var item in collection)
        //        {

        //        }
        //    }
        //}

        //public static IEnumerable<IGrouping<string,RetestUnitModel>> GetSameRetestItemsUnitsCategoryGroupQuery(IEnumerable<IGrouping<string,RetestUnitModel>> unitGroup)
        //{
        //    foreach
        //    IEnumerable<IGrouping<string,RetestUnitModel>> query =  from RetestUnitModel
        //}

        public static int GetRetestItemsCategoryCount(string[] strArray)
        {
            return 0;
        }
    }
}
