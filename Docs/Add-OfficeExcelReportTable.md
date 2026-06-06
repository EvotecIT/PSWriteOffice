---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelReportTable
## SYNOPSIS
Adds an object table to the current Excel report sheet using the OfficeIMO sheet composer.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeExcelReportTable [-Data] <Object[]> [[-Title] <string>] [-TableStyle <string>] [-NoAutoFilter] [-NoFreezeHeaderRow] [-NoAutoFormatDynamicCollections] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds an object table to the current Excel report sheet using the OfficeIMO sheet composer.

## EXAMPLES

### EXAMPLE 1
```powershell
Add-OfficeExcelReportTable -Data @('Value')
```


## PARAMETERS

### -Data
Objects to flatten and render as a table.

```yaml
Type: Object[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoAutoFilter
Disable AutoFilter dropdowns.

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

### -NoAutoFormatDynamicCollections
Disable composer auto-formatting for dynamic collection columns.

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

### -NoFreezeHeaderRow
Do not freeze through the table header row.

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
Emit the A1 range used by the generated table.

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

### -TableStyle
Built-in table style to apply.

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

### -Title
Optional section title displayed above the table.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 1
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
