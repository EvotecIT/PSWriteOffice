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
Add-OfficeExcelReportTable [-InputObject] <Object[]> [[-Title] <string>] [-TableStyle <string>] [-ShowFirstColumn] [-ShowLastColumn] [-NoRowStripes] [-ShowColumnStripes] [-NoAutoFilter] [-NoFreezeHeaderRow] [-NoAutoFormatDynamicCollections] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds an object table to the current Excel report sheet using the OfficeIMO sheet composer.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $rows = @(
    [pscustomobject]@{ Area = 'PDF'; Status = 'Ready' }
    [pscustomobject]@{ Area = 'Word'; Status = 'Review' }
)
New-OfficeExcel -Path .\Operations.xlsx {
    Add-OfficeExcelReportSheet -Name Summary {
        Add-OfficeExcelReportTable -InputObject $rows -Title 'Documentation coverage' -TableStyle TableStyleMedium9
    }
}
```

Renders object rows as a formatted Excel table through the sheet composer.

## PARAMETERS

### -InputObject
Objects to flatten and render as a table.

```yaml
Type: Object[]
Parameter Sets: __AllParameterSets
Aliases: Data
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
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
Accept wildcard characters: False
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
Accept wildcard characters: False
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
Accept wildcard characters: False
```

### -NoRowStripes
Disable alternating row stripes for the generated table.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
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
Accept wildcard characters: False
```

### -ShowColumnStripes
Enable alternating column stripes for the generated table.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowFirstColumn
Emphasize the first table column when the selected style supports it.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowLastColumn
Emphasize the last table column when the selected style supports it.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
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
Accept wildcard characters: False
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
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
