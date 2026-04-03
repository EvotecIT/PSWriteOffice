---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordTable
## SYNOPSIS
Creates a table from PowerShell objects.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordTable [-InputObject] <Object> [[-Content] <scriptblock>] [-Style <WordTableStyle>] [-Layout <string>] [-SkipHeader] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates a table from PowerShell objects.
When used inside `Add-OfficeWordTableCell`, the table is created as a nested table inside that cell.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Add-OfficeWordTable -InputObject $Data -Style 'GridTable1LightAccent1' { WordTableCondition -FilterScript { $_.Total -gt 1000 } }
```

Writes a grid table and highlights rows exceeding $1,000.

### EXAMPLE 2
```powershell
PS>WordTable -Data $Rows {
    WordTableCell -Row 1 -Column 1 {
        WordTable -Data $NestedRows -SkipHeader
    }
}
```

Creates a nested table inside the selected cell.
The same cell scope can also host paragraphs, images, and lists before or after the nested table.

## PARAMETERS

### -Content
DSL content executed inside the table.

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

### -InputObject
Input data (array, list, DataTable, etc.).

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: Data
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Layout
Table layout behavior.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Autofit, Fixed

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the created WordTable.

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

### -SkipHeader
Skip writing header row.

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

### -Style
Built-in table style.

```yaml
Type: WordTableStyle
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: TableNormal, TableGrid, PlainTable1, PlainTable2, PlainTable3, PlainTable4, PlainTable5, GridTable1Light, GridTable1LightAccent1, GridTable1LightAccent2, GridTable1LightAccent3, GridTable1LightAccent4, GridTable1LightAccent5, GridTable1LightAccent6, GridTable2, GridTable2Accent1, GridTable2Accent2, GridTable2Accent3, GridTable2Accent4, GridTable2Accent5, GridTable2Accent6, GridTable3, GridTable3Accent1, GridTable3Accent2, GridTable3Accent3, GridTable3Accent4, GridTable3Accent5, GridTable3Accent6, GridTable4, GridTable4Accent1, GridTable4Accent2, GridTable4Accent3, GridTable4Accent4, GridTable4Accent5, GridTable4Accent6, GridTable5Dark, GridTable5DarkAccent1, GridTable5DarkAccent2, GridTable5DarkAccent3, GridTable5DarkAccent4, GridTable5DarkAccent5, GridTable5DarkAccent6, GridTable6Colorful, GridTable6ColorfulAccent1, GridTable6ColorfulAccent2, GridTable6ColorfulAccent3, GridTable6ColorfulAccent4, GridTable6ColorfulAccent5, GridTable6ColorfulAccent6, GridTable7Colorful, GridTable7ColorfulAccent1, GridTable7ColorfulAccent2, GridTable7ColorfulAccent3, GridTable7ColorfulAccent4, GridTable7ColorfulAccent5, GridTable7ColorfulAccent6, ListTable1Light, ListTable1LightAccent1, ListTable1LightAccent2, ListTable1LightAccent3, ListTable1LightAccent4, ListTable1LightAccent5, ListTable1LightAccent6, ListTable2, ListTable2Accent1, ListTable2Accent2, ListTable2Accent3, ListTable2Accent4, ListTable2Accent5, ListTable2Accent6, ListTable3, ListTable3Accent1, ListTable3Accent2, ListTable3Accent3, ListTable3Accent4, ListTable3Accent5, ListTable3Accent6, ListTable4, ListTable4Accent1, ListTable4Accent2, ListTable4Accent3, ListTable4Accent4, ListTable4Accent5, ListTable4Accent6, ListTable5Dark, ListTable5DarkAccent1, ListTable5DarkAccent2, ListTable5DarkAccent3, ListTable5DarkAccent4, ListTable5DarkAccent5, ListTable5DarkAccent6, ListTable6Colorful, ListTable6ColorfulAccent1, ListTable6ColorfulAccent2, ListTable6ColorfulAccent3, ListTable6ColorfulAccent4, ListTable6ColorfulAccent5, ListTable6ColorfulAccent6, ListTable7Colorful, ListTable7ColorfulAccent1, ListTable7ColorfulAccent2, ListTable7ColorfulAccent3, ListTable7ColorfulAccent4, ListTable7ColorfulAccent5, ListTable7ColorfulAccent6

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
