using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using excel = Microsoft.Office.Interop.Excel;

namespace Bysxiang.UipathExcelEx.Models
{
    public class CellTable : IEnumerable<CellRow>
    {
        private readonly List<RowColumnInfo> cells;

        public ReadOnlyCollection<CellRow> Rows { get; }

        public bool IsEmpty => Rows.Count == 0;

        public CellTable(IList<RowColumnInfo> cells)
        {
            this.cells = new List<RowColumnInfo>(cells);
            var l = (from r in cells.AsParallel() orderby r.CurrentPosition group r by r.CurrentPosition.Row);
            List<CellRow> _rows = new List<CellRow>();
            foreach (var g in l)
            {
                List<RowColumnInfo> items = g.ToList();
                _rows.Add(new CellRow(items));
            }
            Rows = new ReadOnlyCollection<CellRow>(_rows);
        }

        public CellTable(excel.Range regionRng)
        {
            int row = regionRng.Row;
            int maxRow = regionRng.Row + regionRng.Rows.Count - 1;
            int col = regionRng.Column;
            int maxCol = regionRng.Column + regionRng.Columns.Count - 1;

            List<RowColumnInfo> mergeCellList = new List<RowColumnInfo>();
            List<CellRow> _rows = new List<CellRow>();
            foreach (excel.Range rowCell in regionRng.Rows)
            {
                List<RowColumnInfo> _cells = new List<RowColumnInfo>();
                foreach (excel.Range cell in rowCell.Cells)
                {
                    excel.Range mergeArea = cell.MergeArea;
                    if (mergeArea.Row >= row && mergeArea.Row + mergeArea.Rows.Count - 1 <= maxRow
                        && mergeArea.Column >= col && mergeArea.Column + mergeArea.Columns.Count - 1 <= maxCol)
                    {
                        CellPosition p = new CellPosition(cell.Row, cell.Column);
                        if ((bool)cell.MergeCells)
                        {
                            RowColumnInfo c = mergeCellList.Where(c0 => c0.BeginPosition == p).FirstOrDefault();
                            if (c == null)
                            {
                                c = new RowColumnInfo(cell);
                                mergeCellList.Add(c);
                                _cells.Add(c);
                            }
                            else
                            {
                                RowColumnInfo cellInfo = new RowColumnInfo(cell);
                                cellInfo.Value = c.Value;
                                cellInfo.Text = c.Text;
                                _cells.Add(c);
                            }
                        }
                        else
                        {
                            _cells.Add(new RowColumnInfo(cell));
                        }
                    }
                }
                CellRow cellRow = new CellRow(_cells);
                _rows.Add(cellRow);
            }
            Rows = new ReadOnlyCollection<CellRow>(_rows);
            cells = (from r in Rows.AsParallel()
                     from c in r.Items
                     select c).ToList();
        }

        public CellTable this[CellPosition beginPosition, CellPosition endPosition]
        {
            get
            {
                var list = (from r in cells.AsParallel()
                            where r.BeginPosition >= beginPosition && r.EndPosition <= endPosition
                            select r).ToList();
                return new CellTable(list);
            }
        }

        public RowColumnInfo GetRowColumnInfo(CellPosition position)
        {
            return cells.AsParallel().Where(c => c.Container(position)).DefaultIfEmpty(new RowColumnInfo()).First();
        }

        // IEnumerable接口

        public IEnumerator<CellRow> GetEnumerator()
        {
            return Rows.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
