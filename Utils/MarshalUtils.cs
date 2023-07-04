using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Bysxiang.UipathExcelEx.Utils
{
    public static class MarshalUtils
    {
        public static void ReleaseComObject<T>(ref T o) where T : class
        {
            if (o != null && Marshal.IsComObject(o))
            {
                int count = Marshal.ReleaseComObject(o);
            }
        }
    }
}
