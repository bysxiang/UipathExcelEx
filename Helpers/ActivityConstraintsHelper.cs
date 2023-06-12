using System;
using System.Activities;
using System.Activities.Validation;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using UiPath.Excel.Activities;

namespace Bysxiang.UipathExcelEx.Helpers
{
    public static class ActivityConstraintsHelper
    {
        public static Constraint GetCheckParentConstraint<ActivityType>(string parentTypeName, string validationMessage = null) 
            where ActivityType : Activity
        {
            Assembly ass = typeof(ExcelApplicationScope).Assembly;
            Type type = ass.GetType("UiPath.Excel.Activities.CheckParentConstraint");
            if (type != null) // 旧版本
            {
                MethodInfo methodInfo = type.GetMethod("GetCheckParentConstraint",
                    new Type[] { typeof(string), typeof(string) });
                methodInfo = methodInfo.MakeGenericMethod(new Type[] { typeof(ActivityType) });
                return methodInfo.Invoke(null, new object[] { parentTypeName, validationMessage }) as Constraint;
            }
            else
            {
                return GetCheckParentConstraint<ActivityType>(new string[] { parentTypeName }, validationMessage);
            }
        }

        public static Constraint GetCheckParentConstraint<ActivityType>(string[] parentTypeNames, string validationMessage)
            where ActivityType : Activity
        {
            Assembly ass = typeof(ExcelApplicationScope).Assembly;
            Type type = ass.GetType("UiPath.Shared.Activities.ActivityConstraints");
            MethodInfo methodInfo = type.GetMethod("GetCheckParentConstraint", 
                new Type[] { typeof(string[]), typeof(string) });
            methodInfo = methodInfo.MakeGenericMethod(new Type[] { typeof(ActivityType) });
            return methodInfo.Invoke(null, new object[] { parentTypeNames, validationMessage }) as Constraint;
        }
    }
}
