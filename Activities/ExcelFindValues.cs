using Bysxiang.UipathExcelEx.Attributes;
using Bysxiang.UipathExcelEx.Models;
using Bysxiang.UipathExcelEx.Resources;
using Bysxiang.UipathExcelEx.Utils;
using Bysxiang.UipathExcelEx.Views;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using UiPath.Excel;
using UiPath.Excel.Activities;
using excel = Microsoft.Office.Interop.Excel;

namespace Bysxiang.UipathExcelEx.Activities
{
    [LocalDisplayName("ExcelFindValues_Name")]
    [Designer(typeof(ExcelFindValuesView))]
    public class ExcelFindValues : ExcelExInteropActivity<Dictionary<string, RowColumnInfo>>
    {
        [RequiredArgument]
        [LocalizedCategory("Input")]
        [LocalDisplayName("ExcelRange")]
        public InArgument<string> RangeStr { get; set; }

        [RequiredArgument]
        [LocalizedCategory("Input")]
        [LocalDisplayName("ExcelFindValue_Search")]
        public InArgument<IEnumerable<string>> Searchs { get; set; }

        [LocalizedCategory("Input")]
        [LocalDisplayName("ExcelFindValue_MatchFunc")]
        public InArgument<Func<RowColumnInfo, string, bool>> MatchFunc { get; set; }

        [LocalizedCategory("Output")]
        [LocalDisplayName("ExcelResult")]
        public OutArgument<Dictionary<string, RowColumnInfo>> Result { get; set; }

        protected override Task<Dictionary<string, RowColumnInfo>> ExecuteAsync(AsyncCodeActivityContext context, WorkbookApplication workbook)
        {
            string rangeStr = RangeStr.Get(context);
            IEnumerable<string> searchs = Searchs.Get(context).Distinct();
            Func<RowColumnInfo, string, bool> func = MatchFunc.Get(context);

            Dictionary<string, RowColumnInfo> dict = new Dictionary<string, RowColumnInfo>();
            return Task.Run(() =>
            {
                if (string.IsNullOrWhiteSpace(rangeStr))
                {
                    throw new ExcelException(Excel_Activities.ExcelRangeEmptyException);
                }
                if (func == null)
                {
                    func = (cell, s) => cell.IsValid && cell.Value.ToString().Equals(s);
                }

                excel.Worksheet ws = workbook.CurrentWorksheet;
                excel.Range regionRng = null;
                try
                {
                    regionRng = ws.Range[rangeStr];
                }
                catch (COMException ex)
                {
                    throw new ExcelException(string.Format(Excel_Activities.ExcelRangeException, ws.Name, rangeStr), ex);
                }
                foreach (var s in searchs)
                {
                    RowColumnInfo r = ExcelUtils.FindValue(regionRng, null, s, 1, func);
                    dict[s] = r;
                }

                return dict;
            });
        }

        protected override void SetResult(AsyncCodeActivityContext context, Dictionary<string, RowColumnInfo> result)
        {
            this.Result.Set(context, result);
        }
    }
}
