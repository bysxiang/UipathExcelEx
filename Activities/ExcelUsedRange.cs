using System;
using System.Activities;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using UiPath.Excel.Activities;
using up = UiPath.Excel;
using Bysxiang.UipathExcelEx.Helpers;
using Bysxiang.UipathExcelEx.Models;
using excel = Microsoft.Office.Interop.Excel;
using Bysxiang.UipathExcelEx.Attributes;
using Bysxiang.UipathExcelEx.Resources;

namespace Bysxiang.UipathExcelEx.Activities
{
    [LocalDisplayName("ExcelUsedRange_Name")]
    public sealed class ExcelUsedRange : ExcelExInteropActivity<ExcelSizeInfo>
    {
        [LocalizedCategory("Output")]
        public OutArgument<ExcelSizeInfo> SizeInfo { get; set; }

        public ExcelUsedRange():base()
        {
        }

        protected override Task<ExcelSizeInfo> ExecuteAsync(AsyncCodeActivityContext context, up.WorkbookApplication workbook)
        {
            try
            {
                excel.Range range = workbook.CurrentWorksheet.UsedRange;
                return Task.Run<ExcelSizeInfo>(() =>
                {
                    ExcelSizeInfo sizeInfo = new ExcelSizeInfo(range);
                    return sizeInfo;
                });
            }
            catch (COMException ex)
            {
                throw new up.ExcelException(string.Format(Excel_Activities.ExcelUsedRangeException), ex);
            }
        }

        protected override void SetResult(AsyncCodeActivityContext context, ExcelSizeInfo result)
        {
            this.SizeInfo.Set(context, result);
        }
    }
}
