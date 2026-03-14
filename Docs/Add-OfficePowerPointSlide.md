---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePowerPointSlide
## SYNOPSIS
Adds a new slide to a PowerPoint presentation.

## SYNTAX
### Index (Default)
```powershell
Add-OfficePowerPointSlide [[-Content] <scriptblock>] [-Presentation <PowerPointPresentation>] [-Master <int>] [-Layout <int>] [<CommonParameters>]
```

### Name
```powershell
Add-OfficePowerPointSlide [[-Content] <scriptblock>] -LayoutName <string> [-Presentation <PowerPointPresentation>] [-Master <int>] [-CaseSensitive] [<CommonParameters>]
```

### Type
```powershell
Add-OfficePowerPointSlide [[-Content] <scriptblock>] -LayoutType <SlideLayoutValues> [-Presentation <PowerPointPresentation>] [-Master <int>] [<CommonParameters>]
```

## DESCRIPTION
Adds a new slide to a PowerPoint presentation.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>$ppt = New-OfficePowerPoint -FilePath .\deck.pptx; Add-OfficePowerPointSlide -Presentation $ppt
```

Creates a deck and appends a new slide at the end.

### EXAMPLE 2
```powershell
PS>New-OfficePowerPoint -Path .\deck.pptx { PptSlide { PptTitle -Title 'Status Update' } }
```

Creates a slide and sets the title using DSL aliases.

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

### -Content
Nested DSL content executed within the slide scope.

```yaml
Type: ScriptBlock
Parameter Sets: Index, Name, Type
Aliases: None
Possible values: 

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Layout
Layout index to use (matches the template’s built-in layouts).

```yaml
Type: Int32
Parameter Sets: Index
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -LayoutName
Layout name to use (case-insensitive by default).

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

### -Presentation
Presentation to update (optional inside New-OfficePowerPoint).

```yaml
Type: PowerPointPresentation
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

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None

