---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeWordTableCell
## SYNOPSIS
Creates a reusable Word table cell definition for explicit table rows.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeWordTableCell [[-Text] <string>] [-ColumnSpan <int>] [-RowSpan <int>] [<CommonParameters>]
```

## DESCRIPTION
Creates a reusable Word table cell definition for explicit table rows.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $row = @(New-OfficeWordTableCell -Text 'Identity systems' -ColumnSpan 3)
```

The returned cell can be passed to WordTable inside explicit row arrays.

## PARAMETERS

### -ColumnSpan
Number of logical columns covered by the cell.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -RowSpan
Number of logical rows covered by the cell.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Text
Cell text.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `PSWriteOffice.Services.Table.OfficeTableCellSpec` — Describes a logical table cell that can be rendered by multiple Office table surfaces.

## RELATED LINKS

- None
