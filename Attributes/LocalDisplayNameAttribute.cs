using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using r = Bysxiang.UipathExcelEx.Resources;

namespace Bysxiang.UipathExcelEx.Attributes
{
    internal class LocalDisplayNameAttribute : DisplayNameAttribute
    {
        public LocalDisplayNameAttribute(string displayName) : base(displayName)
        {
        }

        public override string DisplayName => r.Resources.ResourceManager.GetString(this.DisplayNameValue) ?? base.DisplayName;
    }
}
