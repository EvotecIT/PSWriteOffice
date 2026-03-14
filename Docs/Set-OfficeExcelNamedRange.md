---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelNamedRange
## SYNOPSIS
Creates or updates a named range.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelNamedRange [-Name] <string> [-Range] <string> [-Hidden] [-ValidationMode <NameValidationMode>] [-Global] [-Save] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelNamedRange [-Name] <string> [-Range] <string> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-Hidden] [-ValidationMode <NameValidationMode>] [-Save] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates or updates a named range.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>ExcelSheet 'Data' { Set-OfficeExcelNamedRange -Name 'Totals' -Range 'B2:B50' }
```

Creates a sheet-scoped name for the range.

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

### -Global
Force a workbook-global name even inside a sheet block.

```yaml
Type: SwitchParameter
Parameter Sets: Context
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Hidden
Mark the defined name as hidden.

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

### -Name
Name of the defined range.

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

### -PassThru
Emit the name after creation.

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
Range in A1 notation.

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

### -Save
Save the workbook immediately after setting the name.

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

### -ValidationMode
Validate or sanitize the defined name.

```yaml
Type: NameValidationMode
Parameter Sets: Context, Document
Aliases: None
Possible values: Sanitize, Strict

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

