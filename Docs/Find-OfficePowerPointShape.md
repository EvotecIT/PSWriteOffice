---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Find-OfficePowerPointShape
## SYNOPSIS
Finds PowerPoint shapes by text, name, kind, or slide.

## SYNTAX
### PresentationText (Default)
```powershell
Find-OfficePowerPointShape [-Text] <string> -Presentation <PowerPointPresentation> [-CaseSensitive] [-Index <int>] [-ShapeIndex <int[]>] [-Name <string[]>] [-Kind <string[]>] [<CommonParameters>]
```

### PresentationRegex
```powershell
Find-OfficePowerPointShape [-Pattern] <string> -Presentation <PowerPointPresentation> [-CaseSensitive] [-Index <int>] [-ShapeIndex <int[]>] [-Name <string[]>] [-Kind <string[]>] [<CommonParameters>]
```

### SlideText
```powershell
Find-OfficePowerPointShape [-Text] <string> -Slide <PowerPointSlide> [-CaseSensitive] [-ShapeIndex <int[]>] [-Name <string[]>] [-Kind <string[]>] [<CommonParameters>]
```

### SlideRegex
```powershell
Find-OfficePowerPointShape [-Pattern] <string> -Slide <PowerPointSlide> [-CaseSensitive] [-ShapeIndex <int[]>] [-Name <string[]>] [-Kind <string[]>] [<CommonParameters>]
```

## DESCRIPTION
Searches an open presentation or a single slide and returns PowerPointShapeInfo
records that include the slide index, shape index, shape kind, extracted text, shape name, and the
underlying OfficeIMO shape object. Text matching includes normal text boxes and table cell text, so
this command can locate the right object before piping it into modification commands.

Use -Text for literal contains matching or -Pattern for regular expressions. Combine
-Kind, -Name, -Index, and -ShapeIndex when a deck has repeated labels and
the script should target a specific slide or shape type.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Find-OfficePowerPointShape -Presentation $ppt -Text 'FY24' -Kind TextBox |
                Set-OfficePowerPointShapeText -Text 'FY25'
```

Finds matching text shapes and updates them without using the PowerPoint DSL.

### EXAMPLE 2
```powershell
PS> $ppt = Get-OfficePowerPoint -Path .\Readiness.pptx
$table = Find-OfficePowerPointShape -Presentation $ppt -Text 'Risk' -Kind Table | Select-Object -First 1
$table | Add-OfficePowerPointTableRow -Values 'Latency', 'Investigating'
$ppt | Close-OfficePowerPoint -Save
```

Uses table-cell text as the locator, then pipes the table shape metadata into a table-row edit.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching for text and name filters.

```yaml
Type: SwitchParameter
Parameter Sets: PresentationText, PresentationRegex, SlideText, SlideRegex
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Index
Optional zero-based slide index when reading from a presentation.

```yaml
Type: Nullable`1
Parameter Sets: PresentationText, PresentationRegex
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Kind
Optional shape kind filter such as TextBox or Table.

```yaml
Type: String[]
Parameter Sets: PresentationText, PresentationRegex, SlideText, SlideRegex
Aliases: None
Possible values: TextBox, Picture, Table, Chart, AutoShape, GroupShape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Name
Optional wildcard filter for shape names.

```yaml
Type: String[]
Parameter Sets: PresentationText, PresentationRegex, SlideText, SlideRegex
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Pattern
Regular expression to find in text boxes and table cells.

```yaml
Type: String
Parameter Sets: PresentationRegex, SlideRegex
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Presentation
Open presentation whose slides should be searched.

```yaml
Type: PowerPointPresentation
Parameter Sets: PresentationText, PresentationRegex
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -ShapeIndex
Optional zero-based shape index filter, useful when several shapes contain the same text.

```yaml
Type: Int32[]
Parameter Sets: PresentationText, PresentationRegex, SlideText, SlideRegex
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Slide
Single slide to search when the caller has already resolved the slide.

```yaml
Type: PowerPointSlide
Parameter Sets: SlideText, SlideRegex
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Text
Literal text to find in text boxes and table cells.

```yaml
Type: String
Parameter Sets: PresentationText, SlideText
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointPresentation
OfficeIMO.PowerPoint.PowerPointSlide`

## OUTPUTS

- `PSWriteOffice.Services.PowerPoint.PowerPointShapeInfo` — PowerShell-friendly description of a PowerPoint shape.

## RELATED LINKS

- None
