using System;
using System.Activities;
using System.Threading.Tasks;
using UiPath.Excel;
using UiPath.Excel.Activities;
using Bysxiang.UipathExcelEx.Helpers;
using UiPath.Excel.Activities.Business;
using UiPath.Excel.Activities.Properties;
using Bysxiang.UipathExcelEx.Attributes;
using Bysxiang.UipathExcelEx.Resources;

namespace Bysxiang.UipathExcelEx.Activities
{
    public abstract class ExcelExInteropActivity<T> : AsyncCodeActivity
    {
        protected bool CreateNew;

        [LocalizedCategory("Input")]
        [LocalDisplayName("ExcelSheet")]
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; } = "Sheet1";

        protected ExcelExInteropActivity()
        {
#if NET461
            this.Constraints.Add(ActivityConstraintsHelper.GetCheckParentConstraint<ExcelExInteropActivity<T>>(new string[2]
            {
                typeof (ExcelApplicationScope).Name,
                typeof (ExcelApplicationCard).Name
            }, string.Format(Excel_Activities.ValidationMessageParents, typeof(ExcelApplicationScope).Name, typeof(ExcelApplicationCard).Name)));
#else
            this.Constraints.Add(ActivityConstraintsHelper.GetCheckParentConstraint<ExcelExInteropActivity<T>>(new string[2]
            {
                typeof (ExcelApplicationScope).Name,
                typeof (ExcelApplicationCard).Name
            }, string.Format(Excel_Activities.ValidationMessageParents, typeof(ExcelApplicationScope).Name, typeof(ExcelApplicationCard).Name)));
#endif
            this.SheetName = new InArgument<string>("Sheet1");
        }
        
        protected override IAsyncResult BeginExecute(AsyncCodeActivityContext context, AsyncCallback callback, 
            object state)
        {
            string sheetName = this.SheetName.Get(context);
            WorkbookApplication workbook = UipathExcelHelper.GetWorkbook(context);
            workbook.SetSheet(sheetName, this.CreateNew);
            Task<T> task = this.ExecuteAsync(context, workbook);
            TaskCompletionSource<T> tacs = new TaskCompletionSource<T>(state);
            Action<Task<T>> continuationAction = (t =>
            {
                workbook.CloseSheet();
                if (t.IsFaulted)
                {
                    tacs.TrySetException(t.Exception.InnerExceptions);
                }    
                else if (t.IsCanceled)
                {
                    tacs.TrySetCanceled();
                }
                else
                {
                    tacs.TrySetResult(t.Result);
                }
                if (callback != null)
                {
                    callback(tacs.Task);
                }
            });
            task.ContinueWith(continuationAction);
            return tacs.Task;
        }

        protected override void EndExecute(AsyncCodeActivityContext context, IAsyncResult result)
        {
            Task<T> task = result as Task<T>;
            if (task.IsFaulted)
                throw task.Exception.InnerException;
            if (!task.IsCanceled)
            {
                if (!context.IsCancellationRequested)
                {
                    try
                    {
                        this.SetResult(context, task.Result);
                        return;
                    }
                    catch (OperationCanceledException)
                    {
                        context.MarkCanceled();
                        return;
                    }
                }
            }
            context.MarkCanceled();
        }

        protected abstract Task<T> ExecuteAsync(AsyncCodeActivityContext context, WorkbookApplication workbook);

        protected abstract void SetResult(AsyncCodeActivityContext context, T result);
    }
}
