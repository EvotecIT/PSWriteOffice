---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePowerPointLayoutPlaceholder
## SYNOPSIS
Gets layout placeholder metadata for a slide.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePowerPointLayoutPlaceholder [-Slide <PowerPointSlide>] [-PlaceholderType <string>] [-Index <uint>] [<CommonParameters>]
```

## DESCRIPTION
Gets layout placeholder metadata for a slide.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficePowerPointLayoutPlaceholder -Slide $slide
```

Returns the layout placeholder definitions for the slide.

### EXAMPLE 2
```powershell
PS>New-OfficePowerPoint -Path .\deck.pptx { PptSlide { Get-OfficePowerPointLayoutPlaceholder } }
```

Uses the current slide context.

## PARAMETERS

### -Index
Optional placeholder index.

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

### -PlaceholderType
Placeholder type to filter on.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: Type
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Slide
Slide to inspect (optional inside DSL).

```yaml
Type: PowerPointSlide
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

- `OfficeIMO.PowerPoint.PowerPointSlide`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointLayoutPlaceholderInfo`

## RELATED LINKS

- None

