using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.Formula.Functions;
using UiPath.Excel.Activities.Business;
using UiPath.Excel.Activities.Properties;
using UiPath.Excel.Activities;
using Bysxiang.UipathExcelEx.Helpers;
using System.Threading.Tasks;
using UiPath.Excel;
using Bysxiang.UipathExcelEx.Resources;

namespace Bysxiang.UipathExcelEx.Activities
{
    public abstract class ExcelExAsyncActivitiy<T> : AsyncCodeActivity
    {
        protected ExcelExAsyncActivitiy()
        {
#if NET461
            this.Constraints.Add(ActivityConstraintsHelper.GetCheckParentConstraint<ExcelExAsyncActivitiy<T>>(new string[2]
            {
                typeof (ExcelApplicationScope).Name,
                typeof (ExcelApplicationCard).Name
            }, string.Format(Excel_Activities.ValidationMessageParents, typeof(ExcelApplicationScope).Name, typeof(ExcelApplicationCard).Name)));
#else
            this.Constraints.Add(ActivityConstraintsHelper.GetCheckParentConstraint<ExcelExAsyncActivitiy<T>>(new string[2]
            {
                typeof (ExcelApplicationScope).Name,
                typeof (ExcelApplicationCard).Name
            }, string.Format(Excel_Activities.ValidationMessageParents, typeof(ExcelApplicationScope).Name, typeof(ExcelApplicationCard).Name)));
#endif
        }

        protected override IAsyncResult BeginExecute(AsyncCodeActivityContext context, AsyncCallback callback, object state)
        {
            WorkbookApplication workbook = UipathExcelHelper.GetWorkbook(context);
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

        protected abstract Task<T> ExecuteAsync(AsyncCodeActivityContext context, WorkbookApplication wba);

        protected abstract void SetResult(AsyncCodeActivityContext context, T result);
    }
}
