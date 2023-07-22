using System;
using System.Activities;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Bysxiang.UipathExcelEx.Attributes;
using Bysxiang.UipathExcelEx.Models;
using Bysxiang.UipathExcelEx.Resources;
using Bysxiang.UipathExcelEx.Utils;
using Bysxiang.UipathExcelEx.Views;
using Microsoft.Office.Interop.Excel;
using UiPath.Excel;
using UiPath.Excel.Activities;
using excel = Microsoft.Office.Interop.Excel;
using r = Bysxiang.UipathExcelEx.Resources;

namespace Bysxiang.UipathExcelEx.Activities
{
    public enum SpecialCellType 
    {
        CellTypeConstants, CellTypeBlanks, CellTypeComments, 
        CellTypeVisible
    }

    [LocalDisplayName("ExcelSpecialCells_Name")]
    [Designer(typeof(ExcelSpecialCellsView))]
    public sealed class ExcelSpecialCells : ExcelExInteropActivity<List<RowColumnInfo>>
    {
        [RequiredArgument]
        [LocalizedCategory("Input")]
        [LocalDisplayName("ExcelRange")]
        public InArgument<string> RangeStr { get; set; }

        [LocalizedCategory("Input")]
        [LocalDisplayName("ExcelSpecialCells_CellType")]
        public SpecialCellType CellType { get; set; } = SpecialCellType.CellTypeConstants;
        
        [Browsable(false)]
        public List<CustomKV<string, SpecialCellType>> TypeList { get; set; }

        [LocalizedCategory("Output")]
        [LocalDisplayName("ExcelSpecialCells_Result")]
        public OutArgument<List<RowColumnInfo>> CellList { get; set; }

        public ExcelSpecialCells() : base()
        {
            var m = r.Resources.ResourceManager;
            List<CustomKV<string, SpecialCellType>> list = new List<CustomKV<string, SpecialCellType>>()
            {
                new CustomKV<string, SpecialCellType>(m.GetString("ExcelSpecialCells#CellTypeConstants"), SpecialCellType.CellTypeConstants),
                new CustomKV<string, SpecialCellType>(m.GetString("ExcelSpecialCells#CellTypeBlanks"), SpecialCellType.CellTypeBlanks),
                new CustomKV<string, SpecialCellType>(m.GetString("ExcelSpecialCells#CellTypeComments"), SpecialCellType.CellTypeComments),
                new CustomKV<string, SpecialCellType>(m.GetString("ExcelSpecialCells#CellTypeVisible"), SpecialCellType.CellTypeVisible)
            };
            this.TypeList = list;
        }

        protected override Task<List<RowColumnInfo>> ExecuteAsync(AsyncCodeActivityContext context, WorkbookApplication workbook)
        {
            string rangeStr = RangeStr.Get(context);
            return Task.Run(() =>
            {
                if (string.IsNullOrWhiteSpace(rangeStr))
                {
                    throw new ExcelException(Excel_Activities.ExcelRangeEmptyException);
                }
                excel.Worksheet ws = workbook.CurrentWorksheet;
                excel.Range regionRng = null;
                try
                {
                    regionRng = ws.Range[rangeStr];
                }
                catch (COMException ex)
                {
                    throw new ExcelException(string.Format(Excel_Activities.ExcelRangeException, ws.Name, rangeStr), ex);
                }

                return ExcelUtils.GetSpecialCellList(regionRng, GetXlCellType(CellType)); 
            });
        }

        protected override void SetResult(AsyncCodeActivityContext context, List<RowColumnInfo> result)
        {
            this.CellList.Set(context, result);
        }
        
        private XlCellType GetXlCellType(SpecialCellType cellType)
        {
            if (cellType == SpecialCellType.CellTypeBlanks)
            {
                return XlCellType.xlCellTypeBlanks;
            }
            else if (cellType == SpecialCellType.CellTypeComments)
            {
                return XlCellType.xlCellTypeComments;
            }
            else if (cellType == SpecialCellType.CellTypeConstants)
            {
                return XlCellType.xlCellTypeConstants;
            }
            else if (cellType == SpecialCellType.CellTypeVisible)
            {
                return XlCellType.xlCellTypeVisible;
            }
            else
            {
                throw new ArgumentException();
            }
        }
    }
}
