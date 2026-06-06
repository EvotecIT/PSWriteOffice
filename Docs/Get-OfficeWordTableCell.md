---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordTableCell
## SYNOPSIS
Gets cells from an OfficeIMO Word table.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeWordTableCell [-Table] <WordTable> [-Row <int>] [-Column <int>] [<CommonParameters>]
```

## DESCRIPTION
Gets cells from an OfficeIMO Word table.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeWordTableCell -Column 'Value'
```


## PARAMETERS

### -Column
Optional zero-based column index.

```yaml
Type: Nullable`1
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Row
Optional zero-based row index.

```yaml
Type: Nullable`1
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Table
Table to inspect.

```yaml
Type: WordTable
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordTable`

## OUTPUTS

- `OfficeIMO.Word.WordTableCell`

## RELATED LINKS

- None
