---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcel
## SYNOPSIS
Opens an existing Excel workbook.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeExcel [-InputPath] <string> [-ReadOnly] [-AutoSave] [<CommonParameters>]
```

## DESCRIPTION
Opens an existing Excel workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>$workbook = Get-OfficeExcel -Path .\report.xlsx -ReadOnly
```

Loads report.xlsx for inspection without enabling writes.

## PARAMETERS

### -AutoSave
Enable automatic saves on the underlying document.

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

### -InputPath
Path to the workbook to load.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: Path, FilePath
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ReadOnly
Open the file in read-only mode.

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

