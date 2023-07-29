using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UiPath.Excel;
using Bysxiang.UipathExcelEx.Attributes;
using Bysxiang.UipathExcelEx.Models;
using excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using UiPath.Excel.Activities;
using System.ComponentModel;
using Bysxiang.UipathExcelEx.Resources;
using Bysxiang.UipathExcelEx.Utils;
using Bysxiang.UipathExcelEx.Views;

namespace Bysxiang.UipathExcelEx.Activities
{
    [LocalDisplayName("ExcelFindValue_Name")]
    [Designer(typeof(FindValueView))]
    public sealed class ExcelFindValue : ExcelExInteropActivity<RowColumnInfo>
    {
        [RequiredArgument]
        [LocalizedCategory("Input")]
        [LocalDisplayName("ExcelRange")]
        public InArgument<string> RangeStr { get; set; }

        [RequiredArgument]
        [LocalizedCategory("Input")]
        [LocalDisplayName("ExcelFindValue_Search")]
        public InArgument<string> Search { get; set; }
        
        [LocalizedCategory("Input")]
        [LocalDisplayName("ExcelFindValue_AfterCell")]
        public InArgument<string> AfterCell { get; set; }

        [RequiredArgument]
        [LocalizedCategory("Input")]
        [LocalDisplayName("ExcelFindValue_WhichNum")]
        public InArgument<int> WhichNum { get; set; } = 1;

        [LocalizedCategory("Input")]
        [LocalDisplayName("ExcelFindValue_MatchFunc")]
        public InArgument<Func<RowColumnInfo, string, bool>> MatchFunc { get; set; }

        [LocalizedCategory("Output")]
        public OutArgument<RowColumnInfo> Result { get; set; }

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);
        }

        protected override Task<RowColumnInfo> ExecuteAsync(AsyncCodeActivityContext context, WorkbookApplication workbook)
        {
            string rangeStr = RangeStr.Get(context);
            string search = Search.Get(context);
            string afterStr = AfterCell.Get(context);
            int whichNum = WhichNum.Get(context);
            Func<RowColumnInfo, string, bool> func = MatchFunc.Get(context);
            if (func == null)
            {
                func = (cell, s) => cell.IsValid && (cell.Value ?? "").ToString().Equals(s);
            }

            return Task.Run(() =>
            {
                if (string.IsNullOrWhiteSpace(rangeStr))
                {
                    throw new ExcelException(Excel_Activities.ExcelRangeEmptyException);
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
                excel.Range afterCell = null;
                if (! string.IsNullOrWhiteSpace(afterStr))
                {
                    try
                    {
                        afterCell = ws.Range[afterStr];
                    }
                    catch (COMException ex)
                    {
                        throw new ExcelException(string.Format(Excel_Activities.ExcelRangeException, ws.Name, afterStr), ex);
                    }
                }
                RowColumnInfo result = ExcelUtils.FindValue(regionRng, afterCell, search, whichNum, func);

                return result;
            });
        }

        protected override void SetResult(AsyncCodeActivityContext context, RowColumnInfo result)
        {
            this.Result.Set(context, result);
        }
    }
}
