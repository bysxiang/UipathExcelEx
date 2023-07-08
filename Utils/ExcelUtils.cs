using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using Bysxiang.UipathExcelEx.Models;
using excel = Microsoft.Office.Interop.Excel;

namespace Bysxiang.UipathExcelEx.Utils
{
    public static class ExcelUtils
    {
        /// <summary>
        /// 列序号转换为列名
        /// </summary>
        /// <param name="colNum"></param>
        /// <returns></returns>
        public static string ToColumnName(int colNum)
        {
            StringBuilder retVal = new StringBuilder();
            int x = 0;

            for (int n = (int)(Math.Log(25 * (colNum + 1)) / Math.Log(26)) - 1; n >= 0; n--)
            {
                x = (int)((Math.Pow(26, (n + 1)) - 1) / 25 - 1);
                if (colNum > x)
                {
                    retVal.Append(Convert.ToChar((int)(((colNum - x - 1) / Math.Pow(26, n)) % 26 + 65)));
                }
            }

            return retVal.ToString();
        }

        /// <summary>
        /// 列名转换为列序号
        /// </summary>
        /// <param name="colName"></param>
        /// <returns></returns>
        public static int ToColumnNum(string colName)
        {
            char[] chars = colName.ToUpper().ToCharArray();

            return (int)(Math.Pow(26, chars.Count() - 1)) *
                (Convert.ToInt32(chars[0]) - 64) +
                ((chars.Count() > 2) ? ToColumnNum(colName.Substring(1, colName.Length - 1)) :
                ((chars.Count() == 2) ? (Convert.ToInt32(chars[chars.Count() - 1]) - 64) : 0));
        }

        /// <summary>
        /// 搜索值，找到第times个值
        /// </summary>
        /// <param name="regionRng">要搜索的区域</param>
        /// <param name="after">从此单元格之后开始搜索</param>
        /// <param name="search">搜索的字符串</param>
        /// <param name="whichNum">第几个</param>
        /// <param name="func">匹配搜索结构的委托</param>
        /// <returns>RowColumnInfo, IsValid标识是否找到了对象</returns>
        public static RowColumnInfo FindValue(excel.Range regionRng, excel.Range after, string search, int whichNum,
            Func<RowColumnInfo, string, bool> func)
        {
            while (whichNum-- > 0)
            {
                RowColumnInfo resultCell = InternalSearchValue(regionRng, after, search, func, out after);
                //Console.WriteLine("after: {0}", after?.Address);
                //if (after != null)
                //{
                //    Marshal.FinalReleaseComObject(after);
                //}
                if (!resultCell.IsValid)
                {
                    return new RowColumnInfo();
                }
                else if (whichNum == 0)
                {
                    return resultCell;
                }
            }

            return new RowColumnInfo();
        }

        /// <summary>
        /// 搜索值，匹配func的对象
        /// </summary>
        /// <param name="regionRng">要搜索的区域</param>
        /// <param name="after">从此单元格之后开始搜索</param>
        /// <param name="search">搜索的字符串</param>
        /// <param name="func">匹配的委托</param>
        /// <param name="resultRng">找到的结果Range对象</param>
        /// <returns>RowColumnInfo, IsValid标识是否找到了对象</returns>
        public static RowColumnInfo InternalSearchValue(excel.Range regionRng, excel.Range after,
            string search, Func<RowColumnInfo, string, bool> func, out excel.Range resultRng)
        {
            excel.Range afterRng = (excel.Range)(after ?? regionRng.Cells[1]);
            RowColumnInfo afterCell = new RowColumnInfo(afterRng);
            if (after == null && func(afterCell, search))
            {
                resultRng = after;
                return afterCell;
            }
            else
            {
                excel.Application app = regionRng.Application;
                excel.CellFormat findFormat = app.FindFormat;
                findFormat?.Clear();
                MarshalUtils.ReleaseComObject(ref findFormat);
                MarshalUtils.ReleaseComObject(ref app);
                
                excel.Range result = null;
                do
                {
                    result = regionRng.Find(What: search, After: afterRng, LookIn: excel.XlFindLookIn.xlValues,
                        LookAt: excel.XlLookAt.xlPart);
                    if (afterRng != null)
                    {
                        //Console.WriteLine("释放1");
                        //MarshalUtils.ReleaseComObject(ref afterRng);
                    }
                    if (result == null)
                    {
                        break;
                    }
                    else
                    {
                        RowColumnInfo resultCell = new RowColumnInfo(result);
                        if (resultCell <= afterCell)
                        {
                            //Console.WriteLine("释放2");
                            //MarshalUtils.ReleaseComObject(ref result);
                            break;
                        }
                        if (resultCell > afterCell && func(resultCell, search))
                        {
                            resultRng = result;
                            return resultCell;
                        }
                        afterRng = result;
                    }
                }
                while (result != null);

                resultRng = null;
                return new RowColumnInfo();
            }
        }

        /// <summary>
        /// 获取一个Range的RowColumnInfo对象，这里仅取独立区域，被合并区域会被忽略
        /// </summary>
        /// <param name="regionRng"></param>
        /// <returns></returns>
        public static List<RowColumnInfo> GetCellList(excel.Range regionRng)
        {
            int row = regionRng.Row;
            int maxRow = regionRng.Row + regionRng.Rows.Count - 1;
            int col = regionRng.Column;
            int maxCol = regionRng.Column + regionRng.Columns.Count - 1;
            List<RowColumnInfo> list = new List<RowColumnInfo>();
            foreach (excel.Range cell in regionRng.Cells)
            {
                excel.Range mergeArea = cell.MergeArea;
                if (mergeArea.Row >= row && mergeArea.Row + mergeArea.Rows.Count - 1 <= maxRow
                    && mergeArea.Column >= col && mergeArea.Column + mergeArea.Columns.Count - 1 <= maxCol)
                {
                    list.Add(new RowColumnInfo(cell));
                }
            }

            return list;
        }
    }
}
