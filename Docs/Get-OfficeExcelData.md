---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelData
## SYNOPSIS
Reads worksheet data as dictionaries or PSCustomObjects.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeExcelData [[-Path] <string>] [-Document <ExcelDocument>] [-Sheet <string>] [-Range <string>] [-NumericAsDecimal] [-AsHashtable] [<CommonParameters>]
```

## DESCRIPTION
Reads worksheet data as dictionaries or PSCustomObjects.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficeExcelData -Path .\report.xlsx -Sheet 'Summary' | Format-Table
```

Returns each row as a PSCustomObject with properties mapped from the header row.

## PARAMETERS

### -AsHashtable
Emit each row as a dictionary instead of PSCustomObjects.

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

### -Document
Workbook to read when it is already open.

```yaml
Type: ExcelDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -NumericAsDecimal
Prefer decimals (instead of doubles) for numeric cells.

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

### -Path
Path to the workbook when no Document is supplied.

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

### -Range
Optional A1 range (e.g. A1:D10). When omitted, the sheet's used range is read.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Sheet
Worksheet name to read; defaults to the first sheet.

```yaml
Type: String
Parameter Sets: __AllParameterSets
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

