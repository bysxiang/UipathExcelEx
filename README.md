# Bysxiang.UipathExcelEx.Activities

## 目录

- [背景](#背景)
- [安装](#安装)
- [用法](#用法)
- [相关项目（可选）](#相关项目)
- [主要项目负责人](#主要项目负责人)
- [参与贡献方式](#参与贡献方式)
    - [贡献人员](#贡献人员)
- [开源协议](#开源协议)

## 背景

由于Uipath自带的Excel组件 —— `Uipath.Excel.Activities`，功能太弱，所以我开发了这个组件库，它扩展了`Uipath.Excel.Activities`，支持`ExcelApplicationScope`和`ExcelApplicationCard`，基于Excel VSTO接口开发。

## 安装

在Uipath中添加`Bysxiang.UipathExcelEx.Activities`安装，它支持`Uipath.Excel.Activities` >= 2.8.6，支持新式Excel与传统Excel模式，支持Windows-旧版(基于.Net Framework4.6.1)与Windows(.Net 5.0以上)。

## 用法

### 模型类说明

ExcelSizeInfo： 描述Excel使用范围情况

- Row 行号
- Column 列号
- RowCount 行数量
- ColumnCount 列数量
- ColumnName 开始列名(如B)
- EndColumnName 尾列列名(如AA)
- FullAddress 使用区域的范围(如 A1:AA5)

WorksheetInfo：描述Sheet

- Name Sheet名称
- Visibility 显示状态
- IsVisible 是否显示
- IsHidden 是否隐藏
- IsVeryHidden 是否是隐藏对象

RowColumnInfo 描述单元格信息

- BeginPosition 合并区域开始坐标
- CurrentPosition 当前开始坐标
- EndPosition 合并区域结束坐标
- Value 单元格的Value值(若单元格无值，将为string.empty，不会为null)
- Text 单元格的Text值(若单元格无值，将为string.empty，不会为null)
- BackgroundColor 背景颜色
- IsValid 是否有效，默认构造函数是一个空对象，它是无效的。
- MergeCells 此单元格是否有合并单元格
- RowCount 合并行数量
- ColCount 合并列数量
- Address 合并行起始单元格地址
- FullAddress 合并区域的地址，如"A1:B1"

CellRow 表示CellTable中的一行，它包括一个`RowColumnInfo`集合。

- IsEmpty 是否不包含任何元素
- Row 第一个元素的`BeginPosition.Row`
- EndRow 第一个元素的`EndPosition.Row`

CellTable 表示一块Excel区域的单元格信息，它包括一个描述所有单元格信息的`RowColumnInfo`集合。为方便操作与查询，被合并区域的Value、Text与合并区域的起始单元格的对应值相同。

例如："A1:D3"是一块合并区域, A1的值是"java"，那么A2、A3的值都是"java"。

- Rows 行集合

从CellTable中返回子对象

	CellTable ct;
	
	ct[beginPosition, endPosition] 从ct中返回在beginPosition和endPositino之间的子对象。

根据坐标，获取一个`RowColumnInfo`对象

	GetRowColumnInfo(CellPosition position)

#### Activities介绍

##### ExcelUsedRange：获取使用信息

- SheetName 工作表名称
- SizeInfo ExcelSizeInfo对象

##### ExcelGetSheetInfos：获取所有Sheet信息

- Sheets Sheet信息集合

##### ExcelSpecialCells 获取指定区域的指定单元格类型

SpecialCellType ： 描述单元格类型的枚举

	public enum SpecialCellType 
	{
	    CellTypeConstants, CellTypeBlanks, CellTypeComments, 
	    CellTypeVisible
	}

- CellType (常量单元格、空单元格、含注释单元格，显示的单元格)
- CellList 获取的结果List，若没有，则返回空集合。仅会包含合并区域的第一个单元格。

##### ExcelFindValue 查找Excel指定范围单元格的值

- RangeStr 查找的范围，不能为空或空字符串
- Search 要搜索的值，支持Excel通配符搜索，如果使用了通配符，则必须提供MatchFunc参数，否则不会找到任何值。
- WhichNum 第几个值，默认为1
- AfterCell 从此之后搜索，不包括这个单元格，例如"A1"
- MatchFunc 搜索匹配委托，一个`Func<RowColumnInfo, string, bool>>`委托，此组件以部分搜索，每次找到单元格都会调用此委托，如果不提供这个参数值，则表示完全匹配。

例如，搜索"java"，默认情况下，"Java"是搜索不到的

##### ExcelReadRange 读取Excel范围

它维护所有行信息，每个行中是`RowColumnInfo`对象，它表示一个单元格，被合并区域也被包括，被合并区域的Value和Text等于合并区域的第一个单元格的值。

- RangeStr 区域
- OutCellTable CellTable对象，单元格合并区域没有完整在其中的不会全部包含。

## 相关项目

暂无

## 主要项目负责人

[@bysxiang](https://github.com/bysxiang)

## 参与贡献方式


### 贡献人员

感谢所有贡献的人。

[@bysxiang](https://github.com/bysxiang)

## 开源协议

[MIT](LICENSE) © bysxiang