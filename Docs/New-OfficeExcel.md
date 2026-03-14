---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeExcel
## SYNOPSIS
Creates a new Excel workbook using the DSL.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeExcel [-FilePath] <string> [[-Content] <scriptblock>] [-AutoSave] [-NoSave] [-Open] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates a new Excel workbook using the DSL.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>New-OfficeExcel -Path .\report.xlsx { ExcelSheet 'Data' { ExcelCell -Address 'A1' -Value 'Region' } }
```

Creates report.xlsx and writes “Region” into cell A1 on the Data worksheet.

## PARAMETERS

### -AutoSave
Opt into OfficeIMO automatic saves during operations.

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

### -Content
DSL scriptblock describing workbook content.

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

### -FilePath
Destination path for the workbook.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: Path
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoSave
Skip saving the workbook after running the DSL.

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

### -Open
Open the workbook in Excel after saving.

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

### -PassThru
Emit a FileInfo for convenience.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None

