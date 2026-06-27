---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Save-OfficeMarkdown
## SYNOPSIS
Saves a Markdown document and optionally creates a PDF sidecar.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficeMarkdown [-Document] <MarkdownDoc> [[-Path] <string>] [-PdfPath <string>] [-WriteOptions <MarkdownWriteOptions>] [-WriteProfile <OfficeMarkdownWriteProfile>] [-ImageRenderingMode <MarkdownImageRenderingMode>] [-LineEnding <string>] [-UnorderedListMarker <string>] [-MarkdownPdfOptions <MarkdownPdfSaveOptions>] [-PdfOptions <PdfOptions>] [-PdfTheme <MarkdownPdfThemeKind>] [-PdfFontFamily <string>] [-PdfTitle <string>] [-PdfAuthor <string>] [-PdfSubject <string>] [-PdfKeywords <string>] [-PdfBaseDirectory <string>] [-PdfApplyWordLikeTheme <bool>] [-PdfIncludeLocalImages <bool>] [-PdfIncludeDataUriImages <bool>] [-PdfRestrictLocalImagesToBaseDirectory <bool>] [-PdfMaximumDataUriImageBytes <int>] [-PdfDefaultImageWidth <double>] [-PdfDefaultImageHeight <double>] [-PdfFrontMatterRenderMode <MarkdownPdfFrontMatterRenderMode>] [-PdfUseFrontMatterVisualTheme <bool>] [-PdfUseFrontMatterMetadata <bool>] [-PdfUseFirstHeadingAsTitle <bool>] [-PdfCreateOutlineFromHeadings <bool>] [-PdfWarningVariable <string>] [-PdfConversionReportVariable <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saves a Markdown document and optionally creates a PDF sidecar.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $doc | Save-OfficeMarkdown -Path .\Report.md -PdfPath .\Report.pdf
```

Writes both artifacts from the same Markdown document model.

## PARAMETERS

### -Document
Markdown document to save.

```yaml
Type: MarkdownDoc
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -ImageRenderingMode
Controls how Markdown images are serialized.

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

### -LineEnding
Markdown line ending: CRLF, LF, CR, or a literal line ending string.

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

### -MarkdownPdfOptions
Advanced Markdown PDF options. Friendly PDF parameters override matching values.

```yaml
Type: MarkdownPdfSaveOptions
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
Emit the Markdown document rather than the saved file.

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

### -Path
Destination Markdown path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PdfApplyWordLikeTheme
Apply the built-in Word-like Markdown PDF baseline theme.

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

### -PdfAuthor
PDF author metadata.

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

### -PdfBaseDirectory
Base directory used to resolve local Markdown images during PDF export.

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

### -PdfConversionReportVariable
Variable name that receives the Markdown PDF conversion report.

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

### -PdfCreateOutlineFromHeadings
Create PDF outlines from Markdown headings.

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

### -PdfDefaultImageHeight
Fallback PDF image height in points.

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

### -PdfDefaultImageWidth
Fallback PDF image width in points.

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

### -PdfFontFamily
Default font family used by Markdown PDF export.

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

### -PdfFrontMatterRenderMode
Controls how YAML front matter appears in the PDF body.

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

### -PdfIncludeDataUriImages
Embed supported data URI images in Markdown PDF output.

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

### -PdfIncludeLocalImages
Embed supported local image files in Markdown PDF output.

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

### -PdfKeywords
PDF keywords metadata.

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

### -PdfMaximumDataUriImageBytes
Maximum decoded bytes for one data URI image in Markdown PDF output.

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

### -PdfOptions
Underlying OfficeIMO.Pdf options used by Markdown PDF export.

```yaml
Type: PdfOptions
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
Optional PDF path to create from the same Markdown document.

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

### -PdfRestrictLocalImagesToBaseDirectory
Require local images to resolve under the base directory.

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

### -PdfSubject
PDF subject metadata.

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

### -PdfTheme
Built-in Markdown PDF visual theme.

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

### -PdfTitle
PDF title metadata.

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

### -PdfUseFirstHeadingAsTitle
Use the first Markdown heading as the PDF title when no title is supplied.

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

### -PdfUseFrontMatterMetadata
Use front matter values as PDF metadata.

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

### -PdfUseFrontMatterVisualTheme
Use front matter values to select a visual theme.

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

### -PdfWarningVariable
Variable name that receives Markdown PDF export warnings.

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

### -UnorderedListMarker
Unordered list marker: '-', '*', or '+'.

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

### -WriteOptions
Optional Markdown writer options.

```yaml
Type: MarkdownWriteOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -WriteProfile
Friendly Markdown writer profile.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDoc
System.IO.FileInfo`

## RELATED LINKS

- None
