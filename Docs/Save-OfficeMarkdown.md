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
Save-OfficeMarkdown [-Document] <MarkdownDoc> [[-Path] <string>] [-PdfPath <string>] [-WriteOptions <MarkdownWriteOptions>] [-WriteProfile <OfficeMarkdownWriteProfile>] [-ImageRenderingMode <MarkdownImageRenderingMode>] [-LineEnding <string>] [-UnorderedListMarker <string>] [-MarkdownPdfOptions <MarkdownPdfSaveOptions>] [-PdfOptions <PdfOptions>] [-PdfTheme <OfficeVisualThemeKind>] [-PdfFontFamily <string>] [-PdfTitle <string>] [-PdfAuthor <string>] [-PdfSubject <string>] [-PdfKeywords <string>] [-PdfBaseDirectory <string>] [-PdfApplyWordLikeTheme <Boolean>] [-PdfIncludeLocalImages <Boolean>] [-PdfIncludeDataUriImages <Boolean>] [-PdfRestrictLocalImagesToBaseDirectory <Boolean>] [-PdfMaximumDataUriImageBytes <Int32>] [-PdfDefaultImageWidth <Double>] [-PdfDefaultImageHeight <Double>] [-PdfFrontMatterRenderMode <MarkdownPdfFrontMatterRenderMode>] [-PdfUseFrontMatterVisualTheme <Boolean>] [-PdfUseFrontMatterMetadata <Boolean>] [-PdfUseFirstHeadingAsTitle <Boolean>] [-PdfCreateOutlineFromHeadings <Boolean>] [-PdfWarningVariable <string>] [-PdfConversionReportVariable <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
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
Accept wildcard characters: False
```

### -ImageRenderingMode
Controls how Markdown images are serialized.

```yaml
Type: MarkdownImageRenderingMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: RichMarkdown, PortableMarkdown, Html

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
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
Accept wildcard characters: False
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
Accept wildcard characters: False
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
Accept wildcard characters: False
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
Accept wildcard characters: False
```

### -PdfApplyWordLikeTheme
Apply the built-in Word-like Markdown PDF baseline theme.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
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
Accept wildcard characters: False
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
Accept wildcard characters: False
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
Accept wildcard characters: False
```

### -PdfCreateOutlineFromHeadings
Create PDF outlines from Markdown headings.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfDefaultImageHeight
Fallback PDF image height in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfDefaultImageWidth
Fallback PDF image width in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
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
Accept wildcard characters: False
```

### -PdfFrontMatterRenderMode
Controls how YAML front matter appears in the PDF body.

```yaml
Type: MarkdownPdfFrontMatterRenderMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Hidden, DocumentHeader, Table

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfIncludeDataUriImages
Embed supported data URI images in Markdown PDF output.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfIncludeLocalImages
Embed supported local image files in Markdown PDF output.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
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
Accept wildcard characters: False
```

### -PdfMaximumDataUriImageBytes
Maximum decoded bytes for one data URI image in Markdown PDF output.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
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
Accept wildcard characters: False
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
Accept wildcard characters: False
```

### -PdfRestrictLocalImagesToBaseDirectory
Require local images to resolve under the base directory.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
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
Accept wildcard characters: False
```

### -PdfTheme
Built-in Markdown PDF visual theme.

```yaml
Type: OfficeVisualThemeKind
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Plain, WordLike, TechnicalDocument, GitHubLike, Compact, Report

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
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
Accept wildcard characters: False
```

### -PdfUseFirstHeadingAsTitle
Use the first Markdown heading as the PDF title when no title is supplied.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfUseFrontMatterMetadata
Use front matter values as PDF metadata.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfUseFrontMatterVisualTheme
Use front matter values to select a visual theme.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
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
Accept wildcard characters: False
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
Accept wildcard characters: False
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
Accept wildcard characters: False
```

### -WriteProfile
Friendly Markdown writer profile.

```yaml
Type: OfficeMarkdownWriteProfile
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: OfficeIMO, Portable, HtmlImage

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDoc`
- `System.IO.FileInfo`

## RELATED LINKS

- None
