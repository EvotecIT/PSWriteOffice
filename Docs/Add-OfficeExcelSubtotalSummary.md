---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelSubtotalSummary
## SYNOPSIS
Adds grouped subtotal summary rows for a worksheet data range.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelSubtotalSummary -GroupColumn <string> -ValueColumn <string[]> [-HeaderRow <int>] [-DataStartRow <int>] [-DataEndRow <int>] [-SummaryStartRow <int>] [-Function <string>] [-NoHeader] [-NoGrandTotal] [-NoOutline] [-HideDetailRows] [-OutlineLevel <int>] [-LabelSuffix <string>] [-GrandTotalLabel <string>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelSubtotalSummary -Document <ExcelDocument> -GroupColumn <string> -ValueColumn <string[]> [-Sheet <string>] [-SheetIndex <int>] [-HeaderRow <int>] [-DataStartRow <int>] [-DataEndRow <int>] [-SummaryStartRow <int>] [-Function <string>] [-NoHeader] [-NoGrandTotal] [-NoOutline] [-HideDetailRows] [-OutlineLevel <int>] [-LabelSuffix <string>] [-GrandTotalLabel <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds grouped subtotal summary rows for a worksheet data range.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet Data { Add-OfficeExcelSubtotalSummary -GroupColumn Region -ValueColumn Sales -DataEndRow 20 }
```

Writes SUBTOTAL formulas below the data range and applies row outline metadata to each group.

## PARAMETERS

### -DataEndRow
Last data row. Defaults to the last row of the used range.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DataStartRow
First data row. Defaults to the row after HeaderRow.

```yaml
Type: Nullable`1
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

### -Function
Subtotal function.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values: Sum, Average, Count, CountNonBlank, Max, Min

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -GrandTotalLabel
Label used for the optional grand total row.

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

### -GroupColumn
Group column as a 1-based index, column letter, or header name.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: By, GroupBy
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HeaderRow
Header row that contains source labels. Defaults to the first row of the used range.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HideDetailRows
Hide detail rows when applying outline metadata.

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

### -LabelSuffix
Text appended to each group key in the subtotal label cell.

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

### -NoGrandTotal
Skip writing a grand total row.

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

### -NoHeader
Skip writing a summary header row.

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

### -NoOutline
Skip applying outline metadata to detail rows.

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

### -OutlineLevel
Outline level used for grouped detail rows.

```yaml
Type: Int32
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
Emit OfficeIMO subtotal generation metadata.

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

### -SummaryStartRow
First row for the generated summary block.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ValueColumn
Value columns as 1-based indexes, column letters, or header names.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: ValueColumns, AggregateColumn, AggregateColumns
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

- `OfficeIMO.Excel.ExcelSubtotalResult`

## RELATED LINKS

- None
