using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ClosedXML;
using DocumentFormat.OpenXml.InkML;
using UiPath.Excel;
using UiPath.Excel.Activities;

namespace Bysxiang.UipathExcelEx.Helpers
{
    internal static class UipathExcelHelper
    {
        /// <summary>
        /// 是否小于2.10.x
        /// </summary>
        /// <returns></returns>
        public static bool IsOldVersion()
        {
            Type t = typeof(ExcelApplicationScope);
            Version v = t.Assembly.GetName().Version;

            return v.MajorRevision == 2 && v.Minor < 10;
        }

        public static WorkbookApplication GetWorkbook(AsyncCodeActivityContext context)
        {
            Assembly ass = typeof(ExcelApplicationScope).Assembly;
            Type t = ass.GetType("UiPath.Excel.Activities.WorkflowDataContextExtensions");
            if (t == null)
            {
                return context.DataContext.GetProperties()[UipathExcelHelper.GetWorkbookScopePropertyTag()].GetValue(context.DataContext) as WorkbookApplication;
            }
            else
            {
                MethodInfo methodInfo = t.GetMethod("GetWorkbookApplication");
                object w = methodInfo.Invoke(null, new object[] {context.DataContext});
                if (w is WorkbookApplication)
                {
                    return w as WorkbookApplication;
                }
                else if (w is WorkbookQuickHandle)
                {
                    IQuickHandleParent parent = (IQuickHandleParent)w;
                    return parent.GetWorkbook() as WorkbookApplication;
                }
                else
                {
                    throw new Exception("无法获取WorkbookApplication");
                }
            }
        }

        public static string GetWorkbookScopePropertyTag()
        {
            Type t = typeof(ExcelApplicationScope);
            FieldInfo field = t.GetField("WorkbookScopePropertyTag", BindingFlags.Static | BindingFlags.NonPublic);
            if (field != null)
            {
                return field.GetValue(t) as string;
            }
            else
            {
                throw new Exception("当前版本不存在WorkbookScopePropertyTag字段");
            }
        }
    }
}
