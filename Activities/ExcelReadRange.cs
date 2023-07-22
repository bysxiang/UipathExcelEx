using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Bysxiang.UipathExcelEx.Activities;
using Bysxiang.UipathExcelEx.Attributes;
using Bysxiang.UipathExcelEx.Models;
using Bysxiang.UipathExcelEx.Resources;
using Bysxiang.UipathExcelEx.Utils;
using Bysxiang.UipathExcelEx.views;
using Microsoft.Office.Interop.Excel;
using UiPath.Excel;
using UiPath.Excel.Activities;
using excel = Microsoft.Office.Interop.Excel;

namespace Bysxiang.UipathExcelEx.Activities
{
    [LocalDisplayName("ExcelReadRange_Name")]
    [Designer(typeof(ReadRangeView))]
    public sealed class ExcelReadRange : ExcelExInteropActivity<CellTable>
    {
        [RequiredArgument]
        [LocalizedCategory("Input")]
        [LocalDisplayName("ExcelRange")]
        public InArgument<string> RangeStr { get; set; }

        [LocalizedCategory("Output")]
        public OutArgument<CellTable> OutCellTable { get; set; }

        protected override Task<CellTable> ExecuteAsync(AsyncCodeActivityContext context, WorkbookApplication workbook)
        {
            string rangeStr = RangeStr.Get(context);
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

                var list = ExcelUtils.GetCellList(regionRng, true);
                var table = new CellTable(list);

                return table;
            });
        }

        protected override void SetResult(AsyncCodeActivityContext context, CellTable result)
        {
            this.OutCellTable.Set(context, result);
        }
    }
}
