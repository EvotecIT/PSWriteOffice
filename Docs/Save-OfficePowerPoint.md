---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Save-OfficePowerPoint
## SYNOPSIS
Saves a presentation to disk.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficePowerPoint -Presentation <PowerPointPresentation> [-Show] [-Password <string>] [-PdfPath <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Invokes the PowerPoint service to persist the document and optionally launch it.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $ppt = New-OfficePowerPoint -FilePath .\Examples\Documents\PowerPointSave.pptx
$slide = Add-OfficePowerPointSlide -Presentation $ppt -Layout 1
Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Saved later'
Save-OfficePowerPoint -Presentation $ppt -PdfPath .\Examples\Documents\PowerPointSave.pdf
```

Saves the current presentation and exports a PDF sidecar.

## PARAMETERS

### -Password
Password used to save the presentation as an encrypted package.

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

### -PdfPath
Optional PDF path to create from the same presentation.

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

### -Presentation
Presentation instance to save.

```yaml
Type: PowerPointPresentation
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Show
Launch the saved file in the default viewer.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None
