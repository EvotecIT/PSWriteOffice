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
```powershell
Copy-OfficePowerPointSlide [[-Presentation] <PowerPointPresentation>] -Index <int> [-InsertAt <int>] [<CommonParameters>]
```

## DESCRIPTION
Duplicates a slide inside the current presentation while preserving shapes, notes, and chart content.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Copy-OfficePowerPointSlide -Presentation $ppt -Index 0
```

Duplicates the first slide and inserts the copy immediately after it.

### EXAMPLE 2
```powershell
PS>Copy-OfficePowerPointSlide -Presentation $ppt -Index 2 -InsertAt 0
```

Duplicates the third slide and inserts the copy at the start of the deck.

## PARAMETERS

### -Presentation
Presentation to update.

```yaml
Type: PowerPointPresentation
Parameter Sets: (All)
Aliases: None
Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Index
Zero-based slide index to duplicate.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: None
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
Parameter Sets: (All)
Aliases: None
Required: False
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

- [Import-OfficePowerPointSlide](Import-OfficePowerPointSlide.md)
