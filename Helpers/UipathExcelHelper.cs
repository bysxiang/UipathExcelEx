using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
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
