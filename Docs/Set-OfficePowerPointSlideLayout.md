---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointSlideLayout
## SYNOPSIS
Changes the layout used by a slide.

## SYNTAX
### Index (Default)
```powershell
Set-OfficePowerPointSlideLayout -Layout <int> [-Slide <PowerPointSlide>] [-Master <int>] [<CommonParameters>]
```

### Name
```powershell
Set-OfficePowerPointSlideLayout -LayoutName <string> [-Slide <PowerPointSlide>] [-Master <int>] [-CaseSensitive] [<CommonParameters>]
```

### Type
```powershell
Set-OfficePowerPointSlideLayout -LayoutType <SlideLayoutValues> [-Slide <PowerPointSlide>] [-Master <int>] [<CommonParameters>]
```

## DESCRIPTION
Changes the layout used by a slide.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Set-OfficePowerPointSlideLayout -LayoutName 'Title and Content'
```

Updates the slide to use the requested layout.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching for layout names.

```yaml
Type: SwitchParameter
Parameter Sets: Name
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Layout
Layout index to use.

```yaml
Type: Int32
Parameter Sets: Index
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -LayoutName
Layout name to use.

```yaml
Type: String
Parameter Sets: Name
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -LayoutType
Layout type to use.

```yaml
Type: SlideLayoutValues
Parameter Sets: Type
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Master
Slide master index to use.

```yaml
Type: Int32
Parameter Sets: Index, Name, Type
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Slide
Slide to update (optional inside a slide DSL scope).

```yaml
Type: PowerPointSlide
Parameter Sets: Index, Name, Type
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

- `OfficeIMO.PowerPoint.PowerPointSlide`

## RELATED LINKS

- None

