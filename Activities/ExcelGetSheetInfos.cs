using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bysxiang.UipathExcelEx.Attributes;
using Bysxiang.UipathExcelEx.Models;
using Bysxiang.UipathExcelEx.Resources;
using Bysxiang.UipathExcelEx.Utils;
using UiPath.Excel;
using UiPath.Excel.Activities;
using UiPath.Shared.Activities;
using excel = Microsoft.Office.Interop.Excel;

namespace Bysxiang.UipathExcelEx.Activities
{
    [LocalDisplayName("ExcelGetSheetInfos_Name")]
    public sealed class ExcelGetSheetInfos : ExcelExAsyncActivitiy<List<WorksheetInfo>>
    {
        [LocalizedCategory("Output")]
        [LocalDisplayName("ExcelResult")]
        public OutArgument<List<WorksheetInfo>> Sheets { get; set; }

        protected override Task<List<WorksheetInfo>> ExecuteAsync(AsyncCodeActivityContext context, WorkbookApplication wba)
        {
            return Task.Run(() =>
            {
                return ExcelUtils.GetSheetList(wba.CurrentWorkbook);
            });
        }

        protected override void SetResult(AsyncCodeActivityContext context, List<WorksheetInfo> result)
        {
            this.Sheets.Set(context, result);
        }
    }
}
