using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using excel = Microsoft.Office.Interop.Excel;

namespace Bysxiang.UipathExcelEx.Models
{
    public enum SheetVisibility
    {
        SheetHidden, SheetVeryHidden, SheetVisible
    }

    public class WorksheetInfo
    {
        public string Name { get; }

        public SheetVisibility Visibility { get; }

        public WorksheetInfo(string _name, excel.XlSheetVisibility _sheetVisibility)
        {
            Name = _name;
            if (_sheetVisibility == excel.XlSheetVisibility.xlSheetHidden)
            {
                Visibility = SheetVisibility.SheetHidden;
            }
            else if (_sheetVisibility == excel.XlSheetVisibility.xlSheetVeryHidden)
            {
                Visibility = SheetVisibility.SheetVeryHidden;
            }
            else if (_sheetVisibility == excel.XlSheetVisibility.xlSheetVisible)
            {
                Visibility = SheetVisibility.SheetVisible;
            }
        }

        public bool IsVisible => Visibility == SheetVisibility.SheetVisible;

        public bool IsHidden => Visibility == SheetVisibility.SheetHidden;

        public bool IsVeryHidden => Visibility == SheetVisibility.SheetVeryHidden;
    }
}
