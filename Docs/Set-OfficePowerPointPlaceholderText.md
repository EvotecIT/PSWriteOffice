---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointPlaceholderText
## SYNOPSIS
Sets text in a slide placeholder.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficePowerPointPlaceholderText -PlaceholderType <string> -Text <string> [-Slide <PowerPointSlide>] [-Index <uint>] [-IgnoreMissing] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets text in a slide placeholder.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Set-OfficePowerPointPlaceholderText -Slide $slide -PlaceholderType Title -Text 'Agenda'
```

Updates the Title placeholder on the slide.

## PARAMETERS

### -IgnoreMissing
Ignore missing placeholders.

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

### -PassThru
Emit the placeholder textbox after update.

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

### -PlaceholderType
Placeholder type to target.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: Type
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Slide
Slide to update (optional inside DSL).

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

### -Text
Text to set.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointTextBox`

## RELATED LINKS

- None

