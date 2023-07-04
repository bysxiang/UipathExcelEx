using System;
using Bysxiang.UipathExcelEx.Utils;

namespace Bysxiang.UipathExcelEx.Models
{
    public class CellPosition : IComparable<CellPosition>, ICloneable
    {
        public int Row { get; set; }

        public int Column { get; set; }

        public CellPosition()
        {
            Row = 0;
            Column = 0;
        }

        public CellPosition(int row, int column)
        {
            Row = row;
            Column = column;
        }

        public CellPosition(string colName, int row)
        {
            Row = row;
            Column = ExcelUtils.ToColumnNum(colName.ToUpper());
        }

        public bool IsValid => Row != 0 && Column != 0;

        public string ExcelRangeName => IsValid ? string.Format("{0}{1}", ExcelUtils.ToColumnName(Column), Row) : "";

        public override bool Equals(object obj)
        {
            return obj is CellPosition position &&
                   Row == position.Row &&
                   Column == position.Column;
        }

        public override int GetHashCode()
        {
            int hashCode = 240067226;
            hashCode = hashCode * -1521134295 + Row.GetHashCode();
            hashCode = hashCode * -1521134295 + Column.GetHashCode();
            return hashCode;
        }

        public int CompareTo(CellPosition other)
        {
            if (Row == other.Row)
            {
                return Column.CompareTo(other.Column);
            }
            else
            {
                return Row.CompareTo(other.Row);
            }
        }

        public object Clone()
        {
            return new CellPosition(Row, Column);
        }

        public static bool operator ==(CellPosition c1, CellPosition c2)
        {
            return c1.Equals(c2);
        }

        public static bool operator !=(CellPosition c1, CellPosition c2)
        {
            return !c1.Equals(c2);
        }

        public static bool operator >(CellPosition c1, CellPosition c2)
        {
            return c1.Row > c2.Row || c1.Column > c2.Column;
        }

        public static bool operator <(CellPosition c1, CellPosition c2)
        {
            return c1.Row < c2.Row || c1.Column < c2.Column;
        }

        public static bool operator >=(CellPosition c1, CellPosition c2)
        {
            return c1 == c2 || c1.Column >= c2.Column;
        }

        public static bool operator <=(CellPosition c1, CellPosition c2)
        {
            return c1 == c2 || c1 < c2;
        }
    }
}
