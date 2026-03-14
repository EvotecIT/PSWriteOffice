---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordTableCondition
## SYNOPSIS
Attaches conditional formatting logic to the current table.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordTableCondition -FilterScript <scriptblock> [-TableStyle <WordTableStyle>] [-BackgroundColor <string>] [<CommonParameters>]
```

## DESCRIPTION
Attaches conditional formatting logic to the current table.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>WordTableCondition -FilterScript { $_.Qty -gt 10 } -BackgroundColor '#fff4d6'
```

Applies a light highlight when the quantity column exceeds 10.

## PARAMETERS

### -BackgroundColor
Row highlight color applied when the predicate matches (ARGB hex).

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

### -FilterScript
Predicate executed per data row (uses $_).

```yaml
Type: ScriptBlock
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -TableStyle
Optional table style applied when the predicate matches.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None

