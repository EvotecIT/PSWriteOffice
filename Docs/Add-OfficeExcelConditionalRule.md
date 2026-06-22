---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelConditionalRule
## SYNOPSIS
Adds a conditional formatting rule to the current worksheet.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelConditionalRule [[-Range] <string>] [[-Operator] <string>] [[-Formula1] <string>] [-HeaderName <string>] [-TableName <string>] [-PivotTableName <string>] [-PivotWholeTable] [-HeaderRow <int>] [-IncludeHeader] [-RuleType <string>] [-Formula2 <string>] [-Text <string>] [-Rank <uint>] [-Percent] [-EqualAverage] [-StandardDeviation <uint>] [-TimePeriod <string>] [-StopIfTrue] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelConditionalRule [[-Range] <string>] [[-Operator] <string>] [[-Formula1] <string>] -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-HeaderName <string>] [-TableName <string>] [-PivotTableName <string>] [-PivotWholeTable] [-HeaderRow <int>] [-IncludeHeader] [-RuleType <string>] [-Formula2 <string>] [-Text <string>] [-Rank <uint>] [-Percent] [-EqualAverage] [-StandardDeviation <uint>] [-TimePeriod <string>] [-StopIfTrue] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a conditional formatting rule to the current worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Add-OfficeExcelConditionalRule -Range 'C2:C100' -Operator GreaterThan -Formula1 '100' }
```

Applies a conditional rule to column C.

## PARAMETERS

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

### -EqualAverage
Include values equal to the average for average rules.

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

### -Formula1
Primary formula or value.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Formula2
Optional secondary formula or value.

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

### -HeaderName
Header or table column name used to resolve the target range.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: ColumnName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HeaderRow
Worksheet header row used when resolving HeaderName without a table. Use 0 for the first row of the used range.

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

### -IncludeHeader
Include the header cell in the resolved range.

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

### -Operator
Conditional formatting operator.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the range after applying the rule.

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

### -Percent
Treat top/bottom rank as a percent.

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

### -PivotTableName
Pivot table name used to resolve the target range.

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

### -PivotWholeTable
Use the full pivot output range instead of the default data body range.

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

### -Range
A1 range to apply the rule to.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Rank
Rank used by top/bottom rules.

```yaml
Type: UInt32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -RuleType
Rule type to author.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: Type
Possible values: CellIs, Expression, Formula, DuplicateValues, UniqueValues, Top, Top10, Bottom, Bottom10, AboveAverage, BelowAverage, ContainsText, NotContainsText, BeginsWith, EndsWith, ContainsBlanks, NotContainsBlanks, ContainsErrors, NotContainsErrors, TimePeriod

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

### -StandardDeviation
Optional standard deviation threshold for average rules.

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

### -StopIfTrue
Stop evaluating later rules when this rule is true.

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

### -TableName
Optional table name for header-based range resolution.

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

### -Text
Text used by text-matching rule types.

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

### -TimePeriod
Time period used by time-period rules.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None
