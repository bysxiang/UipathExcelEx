﻿using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bysxiang.UipathExcelEx.Models
{
    public class RowColumnInfo : IComparable<RowColumnInfo>, ICloneable
    {
        public CellPosition BeginPosition { get; }

        public CellPosition CurrentPosition { get; }

        public CellPosition EndPosition { get; }

        public object Value { get; }

        public string Text { get; }

        public Color BackgroundColor { get; }

        public RowColumnInfo()
        {
            BeginPosition = new CellPosition();
            CurrentPosition = new CellPosition();
            EndPosition = new CellPosition();
            Value = null;
            Text = null;
            BackgroundColor = Color.Black;
        }

        public RowColumnInfo(Range range)
        {
            Range mergeArea = range.MergeArea;
            BeginPosition = new CellPosition(mergeArea.Row, mergeArea.Column);
            CurrentPosition = new CellPosition(range.Row, range.Column);
            EndPosition = new CellPosition(mergeArea.Row + mergeArea.Rows.Count - 1,
                mergeArea.Column + mergeArea.Columns.Count - 1);
            Value = range.Value ?? "";
            Text = range.Text?.ToString() ?? "";
            BackgroundColor = ColorTranslator.FromOle((int)range.DisplayFormat.Interior.Color);
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

        public bool IsValid => BeginPosition.IsValid;

        public bool MergeCells => BeginPosition != EndPosition;

        public int RowCount => EndPosition.Row - BeginPosition.Row + 1;

        public int ColCount => EndPosition.Column - BeginPosition.Column + 1;

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
            return left.BeginPosition.CompareTo(right.BeginPosition) == 0;
        }

        public static bool operator !=(RowColumnInfo left, RowColumnInfo right)
        {
            return !(left == right);
        }
    }
}
