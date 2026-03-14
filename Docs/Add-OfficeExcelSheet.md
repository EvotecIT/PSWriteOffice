---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelSheet
## SYNOPSIS
Adds or reuses a worksheet within the current Excel DSL scope.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeExcelSheet [[-Name] <string>] [[-Content] <scriptblock>] [-ValidationMode <SheetNameValidationMode>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds or reuses a worksheet within the current Excel DSL scope.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>New-OfficeExcel -Path .\report.xlsx { Add-OfficeExcelSheet -Name 'Data' { ExcelCell -Address 'A1' -Value 'Region' } }
```

Creates a workbook with a worksheet named Data and writes the header “Region”.

## PARAMETERS

### -Content
Code to execute inside the worksheet context.

```yaml
Type: ScriptBlock
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Name
Name of the worksheet to create or reuse. When omitted the last sheet is reused or a default sheet is created.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the ExcelSheet object after execution.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ValidationMode
Controls how invalid sheet names are handled.

```yaml
Type: SheetNameValidationMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: None, Sanitize, Strict

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None

