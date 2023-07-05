using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bysxiang.UipathExcelEx.Utils;
using excel = Microsoft.Office.Interop.Excel;
namespace Bysxiang.UipathExcelEx.Models
{
    public class ExcelSizeInfo
    {
        public int Row { get; }

        public int Column { get; }

        public int RowCount { get; }

        public int ColumnCount { get; }

        public string ColumnName => ExcelUtils.ToColumnName(Column);

        public string EndColumnName => ExcelUtils.ToColumnName(Column + ColumnCount - 1);

        public string FullAddress => string.Format("{0}{1}:{2}{3}", ColumnName, Row, EndColumnName, Row + RowCount - 1);

        public ExcelSizeInfo() { }

        public ExcelSizeInfo(excel.Range range)
        {
            Row = range.Row;
            Column = range.Column;
            RowCount = range.Rows.Count;
            ColumnCount = range.Columns.Count;
        }
    }
}
