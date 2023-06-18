using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bysxiang.UipathExcelEx.Models
{
    public class CellTable : IEnumerable<CellRow>
    {
        public ReadOnlyCollection<RowColumnInfo> Cells;

        public ReadOnlyCollection<CellRow> Rows { get; }

        public bool IsEmpty => Rows.Count == 0;

        public CellTable(IList<RowColumnInfo> cells)
        {
            this.Cells = new ReadOnlyCollection<RowColumnInfo>(cells);
            var l = (from r in cells.AsParallel() orderby r.CurrentPosition group r by r.CurrentPosition.Row);
            List<CellRow> _rows = new List<CellRow>();
            foreach (var g in l)
            {
                List<RowColumnInfo> items = g.ToList();
                _rows.Add(new CellRow(items));
            }
            Rows = new ReadOnlyCollection<CellRow>(_rows);
        }

        public CellTable this[CellPosition beginPosition, CellPosition endPosition]
        {
            get
            {
                var list = (from r in Cells.AsParallel()
                            where r.BeginPosition >= beginPosition && r.EndPosition <= endPosition
                            select r).ToList();
                return new CellTable(list);
            }
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
