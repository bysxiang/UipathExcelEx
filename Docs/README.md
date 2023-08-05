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

#### ExcelSizeInfo

描述Excel使用范围情况

> 行号、列号都是从1开始，与Excel中相同

| 属性 | 说明 |
|------| ------ |
| Row |行号，从1开始|
|Column|列号|
|RowCount|行数量|
|ColumnCount| 列数量|
|ColumnName|开始列名|
|EndColumnName| 尾列列名(如AA)|
|FullAddress| 使用区域的范围(如 A1:AA5)|
|DateTimeValue| 将Value转换为DateTime，若转换失败，将抛出`InvalidCastException`，应仅在确认当前单元格是DateTime单元格才执行此操作。

|方法|说明|
|----|----|
|TryGetDateTimeValue| 尝试转换为DateTime|
|ValueEquals|判断值是否匹配|

#### WorksheetInfo

描述Sheet信息

|属性|说明|
|----|-----|
|Name| Sheet名称|
|Visibility| 显示状态|
|IsVisible| 是否显示|
|IsHidden| 是否隐藏|
|IsVeryHidden| 是否是隐藏对象|

#### RowColumnInfo

描述单元格信息

|属性|说明|
|----|----|
|BeginPosition|合并区域开始坐标|
|CurrentPosition |当前开始坐标|
|EndPosition| 合并区域结束坐标|
|Value| 单元格的Value值(若单元格无值，将为string.empty，不会为null)|
|Text |单元格的Text值(若单元格无值，将为string.empty，不会为null)|
|BackgroundColor| 背景颜色|
|IsValid |是否有效，默认构造函数是一个空对象，它是无效的。|
|MergeCells |此单元格是否有合并单元格|
|RowCount |合并行数量|
| ColCount |合并列数量|
| Address |合并行起始单元格地址|
| FullAddress| 合并区域的地址，如"A1:B1"|

#### CellRow

表示CellTable中的一行，它包括一个`RowColumnInfo`集合。

|属性|说明|
|----|----|
|IsEmpty| 是否不包含任何元素|
|Row| 第一个元素的`BeginPosition.Row`|
|EndRow| 第一个元素的`EndPosition.Row`|

#### CellTable

表示一块Excel区域的单元格信息，它包括一个描述所有单元格信息的`RowColumnInfo`集合。为方便操作与查询，被合并区域的Value、Text与合并区域的起始单元格的对应值相同。

例如："A1:D3"是一块合并区域, A1的值是"java"，那么A2、A3的值都是"java"。

属性

|属性|说明|
|----|----|
|Rows| 行集合，`RowColumnInfo`的不可变集合|
|IsEmpty|是否为空|
|this[beginPosition, endPosition]|返回指定坐标区域的子对象|

方法
|方法|说明|
|----|----|
|GetRowColumnInfo|返回坐标处的`RowColumnInfo`对象|

### Activities介绍

#### ExcelUsedRange

获取使用信息，描述可用区域信息

参数说明

|参数|方向|说明|
|----|----|---|
|SheetName|输入|Sheet名称|
|SizeInfo|输出|ExcelSizeInfo对象|

#### ExcelGetSheetInfos

获取所有Sheet信息

|参数|方向|说明|
|----|----|---|
|Sheets|输出|Sheet信息集合|

#### ExcelSpecialCells

获取指定区域的指定单元格列表

SpecialCellType ： 描述单元格类型的枚举

|枚举值|说明|
|------|---|
|CellTypeConstants|常量|
|CellTypeBlanks|空|
|CellTypeComments|含注释单元格|
|CellTypeVisible|显示状态的单元格|

参数

|参数|方向|说明|
|----|----|---|
|CellType|输入|(常量单元格、空单元格、含注释单元格，显示的单元格)|
|CellList|输出| 获取的结果List，若没有，则返回空集合。仅会包含合并区域的第一个单元格。|

#### ExcelFindValue

查找Excel指定范围单元格的值

> 此组件使用VSTO中`Range.Find`进行搜索，因此它支持普通的文本搜索，通配符搜索，搜索日期等等。如果使用了通配符或搜索日期，则必须提供MatchFunc参数，否则无法成功找到要搜索的值。

属性

|参数|方向|说明|
|----|----|---|
|RangeStr|输入| 查找的范围，不能为`null`或空字符串|
|Search|输入| 要搜索的值，支持Excel通配符搜索，如果使用了通配符，则必须提供MatchFunc参数，否则不会找到任何值。|
|WhichNum|输入| 第几个值，默认为1|
|AfterCell|输入| 从此之后搜索，不包括这个单元格，例如"A1"|
|MatchFun|输入| 搜索匹配委托，一个`Func<RowColumnInfo, string, bool>>`委托，此组件以部分搜索，每次找到单元格都会调用此委托，如果不提供这个参数值，则表示完全匹配。|
|Result|Out|找到的`RowColumnInfo`对象，如果未找到，`RowColumnInfo.IsValid`为`false`。|

例如，搜索"java"，默认情况下，"Java"是搜索不到的

#### ExcelReadRange

读取Excel范围，它可以表示合并的单元格信息

它维护所有行信息，每个行中是`RowColumnInfo`对象，它表示一个单元格，被合并区域也被包括，被合并区域的Value和Text等于合并区域的第一个单元格的值。

|参数|方向|说明|
|----|----|----|
|RangeStr|输入|区域|
|OutCellTable|输出|`CellTable`对象，单元格合并区域没有完整在其中的不会全部包含。|

## 相关项目

暂无

## 主要项目负责人

[@bysxiang](https://github.com/bysxiang)

## 参与贡献方式

暂无

### 贡献人员

感谢所有贡献的人。

[@bysxiang](https://github.com/bysxiang)

## 开源协议

[MIT](LICENSE) © bysxiang