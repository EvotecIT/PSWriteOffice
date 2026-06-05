---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficePowerPointDeckPlan
## SYNOPSIS
Creates a semantic PowerPoint deck plan for designer rendering.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficePowerPointDeckPlan [[-Content] <scriptblock>] [<CommonParameters>]
```

## DESCRIPTION
Creates a semantic PowerPoint deck plan for designer rendering.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePowerPointDeckPlan -Content { }
```


## PARAMETERS

### -Content
Nested deck-plan DSL content.

```yaml
Type: ScriptBlock
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

- `OfficeIMO.PowerPoint.PowerPointDeckPlan`

## RELATED LINKS

- None
