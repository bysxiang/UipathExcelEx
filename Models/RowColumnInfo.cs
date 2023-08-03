using System;
using System.Collections.Generic;
using System.Drawing;
using excel = Microsoft.Office.Interop.Excel;

namespace Bysxiang.UipathExcelEx.Models
{
    public class RowColumnInfo : IComparable<RowColumnInfo>, ICloneable
    {
        public CellPosition BeginPosition { get; }

        public CellPosition CurrentPosition { get; }

        public CellPosition EndPosition { get; }

        public object Value { get; set; }

        public DateTime DateTimeValue
        {
            get
            {
                if (Value.GetType() == typeof(DateTime))
                {
                    return (DateTime)Value;
                }
                else if (Value.GetType() == typeof(double))
                {
                    try
                    {
                        return DateTime.FromOADate((double)Value);
                    }
                    catch (OverflowException)
                    {
                        if (DateTime.TryParse(Value.ToString(), out DateTime date))
                        {
                            return date;
                        }
                        else
                        {
                            throw new InvalidCastException();
                        }
                    }
                }
                else
                {
                    throw new InvalidCastException();
                }
            }
        }

        public string Text { get; set; }

        public Color BackgroundColor { get; }

        public RowColumnInfo()
        {
            BeginPosition = new CellPosition();
            CurrentPosition = new CellPosition();
            EndPosition = new CellPosition();
            Value = "";
            Text = "";
            BackgroundColor = Color.Black;
        }

        public RowColumnInfo(excel.Range range)
        {
            excel.Range mergeArea = range.MergeArea;
            excel.Range firstCell = mergeArea.Cells[1] as excel.Range;
            BeginPosition = new CellPosition(mergeArea.Row, mergeArea.Column);
            CurrentPosition = new CellPosition(range.Row, range.Column);
            EndPosition = new CellPosition(mergeArea.Row + mergeArea.Rows.Count - 1,
                mergeArea.Column + mergeArea.Columns.Count - 1);
            Value = firstCell.Value ?? "";
            Text = firstCell.Text?.ToString() ?? "";
            int colorVal = Convert.ToInt32(range.DisplayFormat.Interior.Color);
            BackgroundColor = ColorTranslator.FromOle(colorVal);
        }

        public RowColumnInfo(RowColumnInfo other)
        {
            BeginPosition = (CellPosition)other.BeginPosition.Clone();
            CurrentPosition = (CellPosition)other.CurrentPosition.Clone();
            EndPosition = (CellPosition)other.EndPosition.Clone();
            Value = other.Value;
            Text = other.Text;
            BackgroundColor = other.BackgroundColor;
        }

        public override bool Equals(object obj)
        {
            return obj is RowColumnInfo info &&
                   EqualityComparer<CellPosition>.Default.Equals(BeginPosition, info.BeginPosition) &&
                   EqualityComparer<CellPosition>.Default.Equals(EndPosition, info.EndPosition);
        }

        public override int GetHashCode()
        {
            int hashCode = 913827291;
            hashCode = hashCode * -1521134295 + EqualityComparer<CellPosition>.Default.GetHashCode(BeginPosition);
            hashCode = hashCode * -1521134295 + EqualityComparer<CellPosition>.Default.GetHashCode(EndPosition);
            return hashCode;
        }

        /// <summary>
        /// 是否有效
        /// </summary>
        public bool IsValid => BeginPosition.IsValid;

        public bool MergeCells => BeginPosition != EndPosition;

        public int RowCount => EndPosition.Row - BeginPosition.Row + 1;

        public int ColCount => EndPosition.Column - BeginPosition.Column + 1;

        /// <summary>
        /// 起点地址
        /// </summary>
        public string Address => BeginPosition.ExcelRangeName;

        public string FullAddress
        {
            get
            {
                if (IsValid)
                {
                    if (BeginPosition != EndPosition)
                    {
                        return string.Format("{0}:{1}", BeginPosition.ExcelRangeName, EndPosition.ExcelRangeName);
                    }
                    else
                    {
                        return BeginPosition.ExcelRangeName;
                    }
                }
                else
                {
                    return "";
                }
            }
        }

        /// <summary>
        /// 是否是当前对象的子元素
        /// </summary>
        /// <param name="info"></param>
        /// <returns></returns>
        public bool Container(RowColumnInfo info)
        {
            return Container(info.CurrentPosition) && Container(info.EndPosition);
        }

        /// <summary>
        /// 一个单元格坐标是否包含在当前对象中
        /// </summary>
        /// <param name="position"></param>
        /// <returns></returns>
        public bool Container(CellPosition position)
        {
            return position.IsValid && position >= BeginPosition && position <= EndPosition;
        }

        /// <summary>
        /// 比较开始坐标
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public int CompareTo(RowColumnInfo other)
        {
            return BeginPosition.CompareTo(other.BeginPosition);
        }

        public object Clone()
        {
            return new RowColumnInfo(this);
        }

        public bool TryGetDateTimeValue(out DateTime dateTime)
        {
            if (Value.GetType() == typeof(DateTime))
            {
                dateTime = (DateTime)Value;
                return true;
            }
            else if (Value.GetType() == typeof(double))
            {
                try
                {
                    dateTime = DateTime.FromOADate((double)Value);
                    return true;
                }
                catch (OverflowException)
                {
                    if (DateTime.TryParse(Value.ToString(), out DateTime date))
                    {
                        dateTime = date;
                        return true;
                    }
                    else
                    {
                        dateTime = DateTime.MinValue;
                        return false;
                    }
                }
            }
            else
            {
                dateTime = DateTime.MinValue;
                return false;
            }
        }

        public bool ValueEquals(DateTime dateTime)
        {
            if (this.TryGetDateTimeValue(out DateTime d))
            {
                return d.Equals(dateTime);
            }
            else
            {
                return false;
            }
        }

        public static bool operator >(RowColumnInfo left, RowColumnInfo right)
        {
            return left.BeginPosition.CompareTo(right.BeginPosition) > 0;
        }

        public static bool operator <(RowColumnInfo left, RowColumnInfo right)
        {
            return left.BeginPosition.CompareTo(right.BeginPosition) < 0;
        }

        public static bool operator >=(RowColumnInfo left, RowColumnInfo right)
        {
            return left > right || left == right;
        }

        public static bool operator <=(RowColumnInfo left, RowColumnInfo right)
        {
            return left < right || left == right;
        }

        public static bool operator ==(RowColumnInfo left, RowColumnInfo right)
        {
            if (Object.ReferenceEquals(left, right))
            {
                return true;
            }
            else if (Object.ReferenceEquals(left, null) || Object.ReferenceEquals(right, null))
            {
                return false;
            }
            else
            {
                return left.BeginPosition.CompareTo(right.BeginPosition) == 0;
            }
        }

        public static bool operator !=(RowColumnInfo left, RowColumnInfo right)
        {
            return !(left == right);
        }
    }
}
