---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelPivotTable
## SYNOPSIS
Adds a pivot table to a worksheet.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelPivotTable -SourceRange <string> -DestinationCell <string> [-Name <string>] [-RowField <string[]>] [-ColumnField <string[]>] [-PageField <string[]>] [-DataField <string[]>] [-DataFunction <string[]>] [-DataDisplayName <string[]>] [-DataNumberFormat <string[]>] [-NoRowGrandTotals] [-NoColumnGrandTotals] [-PivotStyle <string>] [-Layout <string>] [-DataOnRows] [-DataOnColumns] [-ShowHeaders] [-HideHeaders] [-ShowEmptyRows] [-HideEmptyRows] [-ShowEmptyColumns] [-HideEmptyColumns] [-ShowDrill] [-HideDrill] [-RowHeaderCaption <string>] [-ColumnHeaderCaption <string>] [-GrandTotalCaption <string>] [-MissingCaption <string>] [-ErrorCaption <string>] [-ShowDataDropDown] [-HideDataDropDown] [-ShowDropZones] [-HideDropZones] [-ShowDataTips] [-HideDataTips] [-ShowMemberPropertyTips] [-HideMemberPropertyTips] [-FieldListSortAscending] [-FieldListSortDescending] [-CustomListSort] [-NoCustomListSort] [-FieldSort <hashtable>] [-FieldHiddenItems <hashtable>] [-FieldVisibleItems <hashtable>] [-PageFieldSelection <hashtable>] [-FieldNoDefaultSubtotal <string[]>] [-FieldSubtotalTop <string[]>] [-FieldInsertBlankRow <string[]>] [-FieldInsertPageBreak <string[]>] [-FieldCompact <string[]>] [-FieldOutline <string[]>] [-FieldHideDropDowns <string[]>] [-RefreshOnOpen] [-NoRefreshOnOpen] [-SaveSourceData] [-NoSaveSourceData] [-PreserveFormatting] [-NoPreserveFormatting] [-EnableDrill] [-DisableDrill] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelPivotTable -Document <ExcelDocument> -SourceRange <string> -DestinationCell <string> [-Sheet <string>] [-SheetIndex <int>] [-Name <string>] [-RowField <string[]>] [-ColumnField <string[]>] [-PageField <string[]>] [-DataField <string[]>] [-DataFunction <string[]>] [-DataDisplayName <string[]>] [-DataNumberFormat <string[]>] [-NoRowGrandTotals] [-NoColumnGrandTotals] [-PivotStyle <string>] [-Layout <string>] [-DataOnRows] [-DataOnColumns] [-ShowHeaders] [-HideHeaders] [-ShowEmptyRows] [-HideEmptyRows] [-ShowEmptyColumns] [-HideEmptyColumns] [-ShowDrill] [-HideDrill] [-RowHeaderCaption <string>] [-ColumnHeaderCaption <string>] [-GrandTotalCaption <string>] [-MissingCaption <string>] [-ErrorCaption <string>] [-ShowDataDropDown] [-HideDataDropDown] [-ShowDropZones] [-HideDropZones] [-ShowDataTips] [-HideDataTips] [-ShowMemberPropertyTips] [-HideMemberPropertyTips] [-FieldListSortAscending] [-FieldListSortDescending] [-CustomListSort] [-NoCustomListSort] [-FieldSort <hashtable>] [-FieldHiddenItems <hashtable>] [-FieldVisibleItems <hashtable>] [-PageFieldSelection <hashtable>] [-FieldNoDefaultSubtotal <string[]>] [-FieldSubtotalTop <string[]>] [-FieldInsertBlankRow <string[]>] [-FieldInsertPageBreak <string[]>] [-FieldCompact <string[]>] [-FieldOutline <string[]>] [-FieldHideDropDowns <string[]>] [-RefreshOnOpen] [-NoRefreshOnOpen] [-SaveSourceData] [-NoSaveSourceData] [-PreserveFormatting] [-NoPreserveFormatting] [-EnableDrill] [-DisableDrill] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a pivot table to a worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $rows = @(
    [pscustomobject]@{ Region = 'North America'; Product = 'Standard'; Sales = 125000 }
    [pscustomobject]@{ Region = 'EMEA'; Product = 'Standard'; Sales = 98000 }
    [pscustomobject]@{ Region = 'APAC'; Product = 'Premium'; Sales = 143000 }
)
New-OfficeExcel -Path .\SalesPivot.xlsx {
    Add-OfficeExcelSheet -Name Data {
        Add-OfficeExcelTable -InputObject $rows -TableName Sales -AutoFit
        Add-OfficeExcelPivotTable -SourceRange 'A1:C4' -DestinationCell 'E2' -Name 'SalesByRegion' -RowField Region -ColumnField Product -DataField Sales -DataFunction Sum -PivotStyle PivotStyleMedium9
    }
}
```

Writes source rows to a worksheet and creates a pivot table using the existing OfficeIMO pivot support.

## PARAMETERS

### -ColumnField
Column fields (header names).

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ColumnHeaderCaption
Optional column header caption.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CustomListSort
Use Excel custom-list sorting.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DataDisplayName
Display names for data fields.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DataField
Data fields (header names). Defaults to the last column when omitted.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DataFunction
Aggregation functions (Sum, Count, Average, etc.).

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DataNumberFormat
Number format codes for data fields.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DataOnColumns
Show data fields on columns.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DataOnRows
Show data fields on rows.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DestinationCell
Top-left destination cell for the pivot table (e.g., "F2").

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DisableDrill
Disable pivot detail drill interaction in Excel.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Workbook to operate on outside the DSL context.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -EnableDrill
Allow users to drill into pivot details in Excel.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ErrorCaption
Optional error-value caption.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FieldCompact
Fields using compact field layout.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FieldHiddenItems
Field item captions to hide, for example @{ Region = @('Legacy') }.

```yaml
Type: Hashtable
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FieldHideDropDowns
Fields whose filter drop-downs should be hidden.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FieldInsertBlankRow
Fields that insert blank rows after items.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FieldInsertPageBreak
Fields that insert page breaks after items.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FieldListSortAscending
Sort pivot field list ascending.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FieldListSortDescending
Sort pivot field list descending.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FieldNoDefaultSubtotal
Fields with default subtotal disabled.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FieldOutline
Fields using outline field layout.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FieldSort
Field sort map, for example @{ Region = 'Ascending' }.

```yaml
Type: Hashtable
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FieldSubtotalTop
Fields with subtotals shown at the top.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FieldVisibleItems
Field item captions to keep visible, hiding other known items.

```yaml
Type: Hashtable
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -GrandTotalCaption
Optional grand total caption.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HideDataDropDown
Hide the data drop-down.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HideDataTips
Hide pivot data tips.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HideDrill
Hide drill indicators.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HideDropZones
Hide pivot drop zones.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HideEmptyColumns
Hide empty columns.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HideEmptyRows
Hide empty rows.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HideHeaders
Hide field headers.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HideMemberPropertyTips
Hide member property tips.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Layout
Pivot layout (Compact, Outline, Tabular).

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MissingCaption
Optional missing-value caption.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Name
Optional pivot table name.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoColumnGrandTotals
Disable column grand totals.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoCustomListSort
Disable Excel custom-list sorting.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoPreserveFormatting
Do not preserve pivot formatting when Excel refreshes the pivot table.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoRefreshOnOpen
Do not refresh the pivot cache when the workbook opens.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoRowGrandTotals
Disable row grand totals.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoSaveSourceData
Do not save pivot source cache records in the workbook package.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PageField
Page fields (header names) used as filters.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PageFieldSelection
Selected page-field item captions, for example @{ Product = 'Standard' }.

```yaml
Type: Hashtable
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the worksheet after creating the pivot table.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PivotStyle
Optional pivot table style name.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PreserveFormatting
Preserve pivot formatting when Excel refreshes the pivot table.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -RefreshOnOpen
Refresh the pivot cache when the workbook opens.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -RowField
Row fields (header names).

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -RowHeaderCaption
Optional row header caption.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SaveSourceData
Save pivot source cache records in the workbook package.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Sheet
Worksheet name when using Document.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SheetIndex
Worksheet index (0-based) when using Document.

```yaml
Type: Nullable`1
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ShowDataDropDown
Show the data drop-down.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ShowDataTips
Show pivot data tips.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ShowDrill
Show drill indicators.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ShowDropZones
Show pivot drop zones.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ShowEmptyColumns
Show empty columns.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ShowEmptyRows
Show empty rows.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ShowHeaders
Show field headers.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ShowMemberPropertyTips
Show member property tips.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SourceRange
Source data range including header row (e.g., "A1:D200").

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None
