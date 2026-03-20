---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertFrom-OfficeWordMarkdown
## SYNOPSIS
Creates a Word document from Markdown.

## SYNTAX
### Markdown (Default)
```powershell
ConvertFrom-OfficeWordMarkdown [-Markdown] <string> [-OutputPath <string>] [-FontFamily <string>] [-BaseUri <string>] [-AllowLocalImages] [-AllowedImageDirectory <string[]>] [-AllowRemoteImages] [-ReaderOptions <MarkdownReaderOptions>] [-FitImagesToPageContentWidth] [-FitImagesToContextWidth] [-MaxImageWidthPixels <double>] [-MaxImageHeightPixels <double>] [-MaxImageWidthPercentOfContent <double>] [-Open] [-PassThru] [<CommonParameters>]
```

### Path
```powershell
ConvertFrom-OfficeWordMarkdown [-FilePath] <string> [-OutputPath <string>] [-FontFamily <string>] [-BaseUri <string>] [-AllowLocalImages] [-AllowedImageDirectory <string[]>] [-AllowRemoteImages] [-ReaderOptions <MarkdownReaderOptions>] [-FitImagesToPageContentWidth] [-FitImagesToContextWidth] [-MaxImageWidthPixels <double>] [-MaxImageHeightPixels <double>] [-MaxImageWidthPercentOfContent <double>] [-Open] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
ConvertFrom-OfficeWordMarkdown -Document <MarkdownDoc> [-OutputPath <string>] [-FontFamily <string>] [-BaseUri <string>] [-AllowLocalImages] [-AllowedImageDirectory <string[]>] [-AllowRemoteImages] [-ReaderOptions <MarkdownReaderOptions>] [-FitImagesToPageContentWidth] [-FitImagesToContextWidth] [-MaxImageWidthPixels <double>] [-MaxImageHeightPixels <double>] [-MaxImageWidthPercentOfContent <double>] [-Open] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates a Word document from Markdown.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>ConvertFrom-OfficeWordMarkdown -Markdown '# Hello' -OutputPath .\hello.docx
```

Writes a Word document containing the supplied Markdown.

### EXAMPLE 2
```powershell
PS>Get-OfficeMarkdown -Path .\README.md | ConvertFrom-OfficeWordMarkdown
```

Returns a Word document instance for further edits.

## PARAMETERS

### -AllowedImageDirectory
Restrict local images to one or more directories.

```yaml
Type: String[]
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AllowLocalImages
Allow local Markdown images to be inserted into the document.

```yaml
Type: SwitchParameter
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AllowRemoteImages
Allow remote HTTP(S) images to be downloaded and inserted.

```yaml
Type: SwitchParameter
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -BaseUri
Base URI used to resolve relative links and images.

```yaml
Type: String
Parameter Sets: Markdown, Path, Document
Aliases: BasePath
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Markdown document instance to convert.

```yaml
Type: MarkdownDoc
Parameter Sets: Document
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -FilePath
Path to a Markdown file.

```yaml
Type: String
Parameter Sets: Path
Aliases: Path
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FitImagesToContextWidth
Fit Markdown images to the current content context width.

```yaml
Type: SwitchParameter
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FitImagesToPageContentWidth
Fit Markdown images to the page content width.

```yaml
Type: SwitchParameter
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FontFamily
Optional font family applied during conversion.

```yaml
Type: String
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Markdown
Markdown text to convert.

```yaml
Type: String
Parameter Sets: Markdown
Aliases: None
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -MaxImageHeightPixels
Optional hard cap for Markdown image height in pixels.

```yaml
Type: Nullable`1
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MaxImageWidthPercentOfContent
Optional hard cap for Markdown image width as a percentage of available content width.

```yaml
Type: Nullable`1
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MaxImageWidthPixels
Optional hard cap for Markdown image width in pixels.

```yaml
Type: Nullable`1
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Open
Open the document after saving.

```yaml
Type: SwitchParameter
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OutputPath
Optional output path for the .docx file.

```yaml
Type: String
Parameter Sets: Markdown, Path, Document
Aliases: OutPath
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit a FileInfo when saving to disk.

```yaml
Type: SwitchParameter
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ReaderOptions
Optional Markdown reader options used before Word conversion.

```yaml
Type: MarkdownReaderOptions
Parameter Sets: Markdown, Path, Document
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

- `System.String
OfficeIMO.Markdown.MarkdownDoc`

## OUTPUTS

- `OfficeIMO.Word.WordDocument
System.IO.FileInfo`

## RELATED LINKS

- None

