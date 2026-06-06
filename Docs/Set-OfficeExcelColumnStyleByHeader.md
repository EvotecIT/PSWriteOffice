---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelColumnStyleByHeader
## SYNOPSIS
Applies common number, fill, font, and status styles to a worksheet column resolved by header text.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelColumnStyleByHeader [-Header] <string> [-IncludeHeader] [-Style <string>] [-Decimals <int>] [-CultureName <string>] [-NumberFormat <string>] [-Pattern <string>] [-Bold] [-BackgroundColor <string>] [-FontColor <string>] [-Alignment <string>] [-BackgroundByText <hashtable>] [-FontColorByText <hashtable>] [-BoldByText <string[]>] [-CaseSensitive] [-Width <double>] [-AutoFit] [-IgnoreMissing] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelColumnStyleByHeader [-Header] <string> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-IncludeHeader] [-Style <string>] [-Decimals <int>] [-CultureName <string>] [-NumberFormat <string>] [-Pattern <string>] [-Bold] [-BackgroundColor <string>] [-FontColor <string>] [-Alignment <string>] [-BackgroundByText <hashtable>] [-FontColorByText <hashtable>] [-BoldByText <string[]>] [-CaseSensitive] [-Width <double>] [-AutoFit] [-IgnoreMissing] [<CommonParameters>]
```

## DESCRIPTION
Uses the OfficeIMO header resolver so scripts can style report columns without calculating column letters or ranges.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' {
                Set-OfficeExcelColumnStyleByHeader -Header Revenue -Style Currency -CultureName en-US -AutoFit
                Set-OfficeExcelColumnStyleByHeader -Header Status -BackgroundByText @{ Ready = '#D4EDDA'; Blocked = '#F8D7DA' } -BoldByText Blocked
              }
```

Styles Revenue as currency and colors Status cells by their text.

## PARAMETERS

### -Alignment
Align cell content in the resolved column.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values: Left, Center, Right

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AutoFit
Auto-fit the resolved column after applying styles.

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

### -BackgroundByText
Background colors keyed by matching cell text.

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

### -BackgroundColor
Apply a solid background color to the whole resolved column.

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

### -Bold
Apply bold text to the whole resolved column.

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

### -BoldByText
Values that should be bolded when the cell text matches.

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

### -CaseSensitive
Use case-sensitive matching for text maps.

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

### -CultureName
Culture used by currency formatting, such as en-US or pl-PL.

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

### -Decimals
Decimal places for number, percent, and currency styles.

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

### -FontColor
Apply a font color to the whole resolved column.

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

### -FontColorByText
Font colors keyed by matching cell text.

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

### -Header
Header caption used to resolve the target column.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -IgnoreMissing
Do nothing when the header cannot be found instead of throwing.

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

### -IncludeHeader
Include the header cell in the applied formatting.

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

### -NumberFormat
Custom number format. Also used when Style is NumberFormat.

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

### -Pattern
Date or DateTime number format pattern.

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

### -Style
Preset number style to apply.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values: Number, Integer, Percent, Currency, Date, DateTime, Time, DurationHours, Text, NumberFormat

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Width
Set the resolved column width.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None
