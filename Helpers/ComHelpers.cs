using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Bysxiang.UipathExcelEx.Helpers
{
    public static class ComHelpers
    {
        public static void ReleaseAndClearComObject<T>(ref T o) where T : class
        {
            if (o != null)
            {
                if (Marshal.IsComObject(o))
                {
                    Marshal.ReleaseComObject(o);
                }
                o = default(T);
            }
        }
    }
}
