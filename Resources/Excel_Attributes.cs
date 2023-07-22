using System;
using System.Collections.Generic;
using System.Linq;
using System.Resources;
using System.Text;
using System.Threading.Tasks;

namespace Bysxiang.UipathExcelEx.Resources
{
    public static class Excel_Activities
    {
        private static ResourceManager rm = Resources.ResourceManager;

        public static string ExcelSheet => rm.GetString(nameof(ExcelSheet));

        public static string ExcelRange => rm.GetString(nameof(ExcelRange));

        public static string ExcelUsedRange_Name => rm.GetString(nameof(ExcelUsedRange_Name));

        public static string ExcelUsedRangeException => rm.GetString(nameof(ExcelUsedRangeException));

        public static string ExcelFindValue_Name => rm.GetString(nameof(ExcelFindValue_Name));

        public static string ExcelFindValue_Search => rm.GetString(nameof(ExcelFindValue_Search));

        public static string ExcelFindValue_WhichNum => rm.GetString(nameof(ExcelFindValue_WhichNum));

        public static string ExcelRangeException => rm.GetString(nameof(ExcelRangeException));

        public static string ExcelRangeEmptyException => rm.GetString(nameof(ExcelRangeEmptyException));

        public static string ExcelReadRange_Name => rm.GetString(nameof(ExcelReadRange_Name));
    }
}
