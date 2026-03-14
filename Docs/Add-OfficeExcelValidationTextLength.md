---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelValidationTextLength
## SYNOPSIS
Adds a text-length data validation rule to a worksheet range.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelValidationTextLength [-Range] <string> [-Operator] <string> [-Formula1] <int> [-Formula2 <int>] [-AllowBlank <bool>] [-ErrorTitle <string>] [-ErrorMessage <string>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelValidationTextLength [-Range] <string> [-Operator] <string> [-Formula1] <int> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-Formula2 <int>] [-AllowBlank <bool>] [-ErrorTitle <string>] [-ErrorMessage <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a text-length data validation rule to a worksheet range.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>ExcelSheet 'Data' { Add-OfficeExcelValidationTextLength -Range 'E2:E20' -Operator Between -Formula1 1 -Formula2 50 }
```

Ensures text length in E2:E20 is between 1 and 50 characters.

## PARAMETERS

### -AllowBlank
Allow blank values.

```yaml
Type: Boolean
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

### -ErrorMessage
Error message shown to the user.

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

### -ErrorTitle
Error title shown to the user.

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

### -Formula1
Primary length bound.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Formula2
Optional secondary length bound.

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

### -Operator
Validation operator.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the range string after applying validation.

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
Target range in A1 notation.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None

