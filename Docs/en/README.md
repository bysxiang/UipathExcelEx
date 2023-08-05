# Bysxiang.UipathExcelEx.Activities

## 目录

- [background](#background)
- [Install](#Install)
- [Usage](#Usage)
- [Related projects（optional）](#Related%20projects)
- [Main project manager](#Main%20project%20manager)
- [Mode of participation and contribution](#Mode%20of%20participation%20and%20contribution)
    - [Contributor](#Contributor)
- [Open source agreement](#Open%source%agreement)

## background

Because the Excel component of Uipath-`Uipath.Excel.Activities` is too weak, I developed this component library, which extends `Uipath.Excel.Activities`, supports `ExcelApplicationScope` and `ExcelApplicationCard`, and is developed based on Excel VSTO interface.

## Install

Add `Bysxiang.UipathExcelEx.Activities` installation to Uipath, which supports `Uipath.Excel.Activities` > = 2.8.6, new Excel and traditional Excel mode, old Windows- (based on .net Framework4.6.1) and Windows (.net 5.0 +)。

## Usage

### Model class description

#### ExcelSizeInfo

Describe the scope of Excel usage

> The line number and column number all start at 1, the same as in Excel。

| Attribute | Description |
|------| ------ |
| Row |Line number, starting with 1|
|Column|Column number|
|RowCount|Row Count|
|ColumnCount| Column Count|
|ColumnName|Start the column name|
|EndColumnName| End column name(For example "AA")|
|FullAddress| The scope of the area of use(For example "A1:AA5")|
|DateTimeValue| Convert Value to DateTime. If the conversion fails, `InvalidCastException` will be thrown. You should do this only if you confirm that the current cell is a DateTime cell.

|Method|Description|
|----|----|
|TryGetDateTimeValue| Attempt to convert to DateTime|
|ValueEquals|Determine whether the values match|

#### WorksheetInfo

Describe Sheet information

| Attribute | Description |
|----|-----|
|Name| Sheet Name|
|Visibility| Visble status|
|IsVisible| Is Visible|
|IsHidden| Is Hidden|
|IsVeryHidden| Whether it is a hidden object|

#### RowColumnInfo

Describe cell information

| Attribute | Description |
|----|----|
| BeginPosition | start coordinates of the merge area |
| CurrentPosition | current start coordinates |
| EndPosition | end coordinates of the merge area |
| Value | value of the cell (if the cell has no value, it will be string.empty, not null) |
| Text | text value of the cell (if the cell has no value, it will be string.empty, not null) |
| BackgroundColor | background color |
| IsValid | whether it is valid or not. The default constructor is an empty object, which is invalid. |
| MergeCells | whether this cell has merged cells |
| RowCount | number of merged rows |
| ColCount | number of juxtaposition |
| Address | start cell address of the merged row |
| FullAddress | address of the merge zone, such as "A1:B1" |

#### CellRow

Represents a row in CellTable that includes a collection of `RowColumnInfo`.

| Attribute | Description |
|----|----|
| IsEmpty | whether it contains no elements |
| Row | `BeginPosition.Row` of the first element |
| EndRow | `EndPosition.Row` of the first element |

#### CellTable

Represents the cell information of an Excel range, which includes a collection of `RowColumnInfo` that describes all cell information. To facilitate operation and query, the Value and Text of the merged region are the same as the corresponding values of the starting cell of the merged region.

For example: "A1:D3" is a merged area, the value of A1 is "java", then the values of "A2" and "A3" are "java".

Attribute

| Attribute | Description |
|----|----|
|Rows| 行集合，`RowColumnInfo`的不可变集合|
|IsEmpty|是否为空|
|this[beginPosition, endPosition]|返回指定坐标区域的子对象|
| Rows | Row collection, immutable collection of `RowColumnInfo` |
| IsEmpty | whether it is empty |
| this [beginPosition, endPosition] | returns sub-objects in the specified coordinate area |

Method

|Method|Description|
|----|----|
|GetRowColumnInfo|Returns the `RowColumnInfo` object at the coordinates|

### Activities introduction

#### ExcelUsedRange

Get usage information and describe the available area information. 
Parameter description

| Parameter | Direction | description |
|----|----|---|
|SheetName|In|Sheet Name|
|SizeInfo|Out|`ExcelSizeInfo` Object|

#### ExcelGetSheetInfos

Get All Sheet Infomation

| Parameter | Direction | description |
|----|----|---|
|Sheets|Out|`SheetInfo` collection|

#### ExcelSpecialCells

Gets the list of specified cells for the specified area

SpecialCellType ： Enumerations that describe cell types

|Enum|Description|
|------|---|
| CellTypeConstants | constant |
| CellTypeBlanks | Null |
| CellTypeComments | with comment cell |
| CellTypeVisible | display status cell |

Argument

| Parameter | Direction | description |
|----|----|---|
| CellType | input | (constant cell, empty cell, annotated cell, displayed cell) |
| CellList | output | List of the obtained result. If not, an empty collection is returned. Only the first cell of the merged range will be included. |

#### ExcelFindValue

Find the value of the cell in the range specified by Excel.

> This component uses `Range.Find` in `VSTO` to search, so it supports normal text search, wildcard search, search date, and so on. If you use wildcards or search dates, you must provide the `MatchFunc` parameter, or you cannot successfully find the value you want to search for.

Parameter

| Parameter | Direction | description |
|----|----|---|
|RangeStr|In| The scope of the lookup, which cannot be `null` or empty string|
|Search|In| The value to search for supports Excel wildcard search, and if wildcards are used, the `MatchFunc` parameter must be provided, otherwise no value will be found.|
|WhichNum|In| The number of values. Default is 1.|
|AfterCell|In| Search from then on, excluding this cell, such as "A1"|
|MatchFun|In| Search match delegate, a `Func<RowColumnInfo, string, bool>` delegate, this component searches in part, and this delegate will be called every time a cell is found, if this parameter value is not provided, it means a complete match.|
|Result|Out|The found `RowColumnInfo` object, if not found, `RowColumnInfo.IsValid` is `false`.|

For example, search for "java", by default, "Java" is not searchable.

#### ExcelReadRange

Read the Excel range, which can represent merged cell information.

It maintains all row information, and in each row is the `RowColumnInfo` object, which represents a cell, the merged region is also included, and the Value and Text of the merged region are equal to the value of the first cell of the merged region.

| Parameter | Direction | description |
|----|----|----|
|RangeStr|In|Excel Range Str|
|OutCellTable|Out|`CellTable` object, all of which are not fully included in the cell merge range.|

## Related projects

None for the time being

## Main project manager

[@bysxiang](https://github.com/bysxiang)

## Mode of participation and contribution

None for the time being

### Contributor

感谢所有贡献的人。

[@bysxiang](https://github.com/bysxiang)

## Open source agreement

[MIT](LICENSE) © bysxiang