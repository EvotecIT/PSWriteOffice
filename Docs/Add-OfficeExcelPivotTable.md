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
Add-OfficeExcelPivotTable -SourceRange <string> -DestinationCell <string> [-Name <string>] [-RowField <string[]>] [-ColumnField <string[]>] [-PageField <string[]>] [-DataField <string[]>] [-DataFunction <string[]>] [-NoRowGrandTotals] [-NoColumnGrandTotals] [-PivotStyle <string>] [-Layout <string>] [-DataOnRows] [-DataOnColumns] [-ShowHeaders] [-HideHeaders] [-ShowEmptyRows] [-HideEmptyRows] [-ShowEmptyColumns] [-HideEmptyColumns] [-ShowDrill] [-HideDrill] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelPivotTable -Document <ExcelDocument> -SourceRange <string> -DestinationCell <string> [-Sheet <string>] [-SheetIndex <int>] [-Name <string>] [-RowField <string[]>] [-ColumnField <string[]>] [-PageField <string[]>] [-DataField <string[]>] [-DataFunction <string[]>] [-NoRowGrandTotals] [-NoColumnGrandTotals] [-PivotStyle <string>] [-Layout <string>] [-DataOnRows] [-DataOnColumns] [-ShowHeaders] [-HideHeaders] [-ShowEmptyRows] [-HideEmptyRows] [-ShowEmptyColumns] [-HideEmptyColumns] [-ShowDrill] [-HideDrill] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a pivot table to a worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Add-OfficeExcelPivotTable -SourceRange 'A1:D200' -DestinationCell 'F2' -RowField 'Region' -DataField 'Sales'
```

Creates a pivot table in F2 using Region as rows and Sales as the data field.

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

