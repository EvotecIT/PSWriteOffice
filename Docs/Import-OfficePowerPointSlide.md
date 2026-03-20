---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Import-OfficePowerPointSlide
## SYNOPSIS
Imports a slide from another PowerPoint presentation.

## SYNTAX
### SourcePresentation (Default)
```powershell
Import-OfficePowerPointSlide -SourcePresentation <PowerPointPresentation> -SourceIndex <int> [-Presentation <PowerPointPresentation>] [-InsertAt <int>] [<CommonParameters>]
```

### SourcePath
```powershell
Import-OfficePowerPointSlide -SourcePath <string> -SourceIndex <int> [-Presentation <PowerPointPresentation>] [-InsertAt <int>] [<CommonParameters>]
```

## DESCRIPTION
Imports a slide from another PowerPoint presentation.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Import-OfficePowerPointSlide -Presentation $target -SourcePath .\source.pptx -SourceIndex 0
```

Copies the first slide from source.pptx into the target presentation.

## PARAMETERS

### -InsertAt
Optional target insertion index; omit to append.

```yaml
Type: Nullable`1
Parameter Sets: SourcePresentation, SourcePath
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Presentation
Target presentation to update (optional inside DSL).

```yaml
Type: PowerPointPresentation
Parameter Sets: SourcePresentation, SourcePath
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -SourceIndex
Zero-based slide index in the source presentation.

```yaml
Type: Int32
Parameter Sets: SourcePresentation, SourcePath
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SourcePath
Path to the source presentation.

```yaml
Type: String
Parameter Sets: SourcePath
Aliases: Path
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SourcePresentation
Source presentation to import from.

```yaml
Type: PowerPointPresentation
Parameter Sets: SourcePresentation
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

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## RELATED LINKS

- None

