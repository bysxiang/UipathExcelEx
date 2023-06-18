using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bysxiang.UipathExcelEx.Models
{
    public class CellRow : IEnumerable<RowColumnInfo>
    {
        public ReadOnlyCollection<RowColumnInfo> Items { get; }

        public CellRow(IList<RowColumnInfo> cells)
        {
            this.Items = new ReadOnlyCollection<RowColumnInfo>(cells);
        }

        public bool IsEmpty => Items.Count == 0;

        public int Row => Items.Count > 0 ? Items.First().BeginPosition.Row : 0;

        public int EndRow => Items.Count > 0 ? Items.First().EndPosition.Row : 0;

        public List<RowColumnInfo> GetItems(int startColumn, int endColumn)
        {
            return (from i in Items
                    where i.BeginPosition.Column >= startColumn && i.EndPosition.Column <= endColumn
                    select i).ToList();
        }

        public IEnumerator<RowColumnInfo> GetEnumerator()
        {
            return Items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
