---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficeWordImage
## SYNOPSIS
Exports a Word page as PNG or SVG with structured image diagnostics.

## SYNTAX
### Path (Default)
```powershell
Export-OfficeWordImage [-Path] <string> [-OutputPath] <string> [-Format <OfficeImageExportFormat>] [-Options <WordImageExportOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Export-OfficeWordImage [-OutputPath] <string> -Document <WordDocument> [-Format <OfficeImageExportFormat>] [-Options <WordImageExportOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Exports a Word page as PNG or SVG with structured image diagnostics.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Export-OfficeWordImage -Path .\Report.docx -OutputPath .\Report.svg -Format Svg
```

Returns the OfficeIMO image export result after writing the image.

## PARAMETERS

### -Document
Open Word document instance.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Format
Output image format.

```yaml
Type: OfficeImageExportFormat
Parameter Sets: Path, Document
Aliases: None
Possible values: Png, Svg, Jpeg, Tiff, Webp

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Options
Optional page, size, scale, theme, and rendering settings.

```yaml
Type: WordImageExportOptions
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OutputPath
Destination PNG or SVG path.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
Path to the Word document.

```yaml
Type: String
Parameter Sets: Path
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

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Drawing.OfficeImageExportResult`

## RELATED LINKS

- None
