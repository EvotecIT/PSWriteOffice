---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Copy-OfficePowerPointSlide
## SYNOPSIS
Copies an existing slide within a PowerPoint presentation.

## SYNTAX
### __AllParameterSets
```powershell
Copy-OfficePowerPointSlide -Index <int> [-Presentation <PowerPointPresentation>] [-InsertAt <int>] [<CommonParameters>]
```

## DESCRIPTION
Copies an existing slide within a PowerPoint presentation.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Copy-OfficePowerPointSlide -Presentation $ppt -Index 0
```

Creates a duplicate of slide 1 and inserts it at position 2.

## PARAMETERS

### -Index
Zero-based slide index to duplicate.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -InsertAt
Optional target index for the duplicate; omit to insert after the source slide.

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

### -Presentation
Presentation to update (optional inside DSL).

```yaml
Type: PowerPointPresentation
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## RELATED LINKS

- None

