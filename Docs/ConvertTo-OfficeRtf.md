---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficeRtf
## SYNOPSIS
Converts Word, HTML, or PDF input to RTF.

## SYNTAX
### WordPath (Default)
```powershell
ConvertTo-OfficeRtf -WordPath <string> [-OutputPath <string>] [-PassThru] [<CommonParameters>]
```

### WordDocument
```powershell
ConvertTo-OfficeRtf -WordDocument <WordDocument> [-OutputPath <string>] [-PassThru] [<CommonParameters>]
```

### Html
```powershell
ConvertTo-OfficeRtf -Html <string> [-OutputPath <string>] [-FontFamily <string>] [-BasePath <string>] [-StylesheetPath <string[]>] [-StylesheetContent <string[]>] [-PassThru] [<CommonParameters>]
```

### HtmlPath
```powershell
ConvertTo-OfficeRtf -HtmlPath <string> [-OutputPath <string>] [-FontFamily <string>] [-BasePath <string>] [-StylesheetPath <string[]>] [-StylesheetContent <string[]>] [-PassThru] [<CommonParameters>]
```

### PdfPath
```powershell
ConvertTo-OfficeRtf -PdfPath <string> [-OutputPath <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Converts Word, HTML, or PDF input to RTF.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\Report.docx { WordParagraph -Text 'Summary' }
            ConvertTo-OfficeRtf -WordPath .\Report.docx -OutputPath .\Report.rtf -PassThru
```

Loads the Word document and saves an RTF file using OfficeIMO.Word.Rtf.

### EXAMPLE 2
```powershell
PS> ConvertTo-OfficeRtf -Html '<h1>Report</h1>' -OutputPath .\Report.rtf
```

Creates a Word document from HTML and serializes it to RTF.

### EXAMPLE 3
```powershell
PS> ConvertTo-OfficeRtf -PdfPath .\Report.pdf -OutputPath .\Report.rtf
```

Uses OfficeIMO.Rtf.Pdf's semantic PDF reader to write RTF output.

## PARAMETERS

### -BasePath
Base path used to resolve relative HTML resources.

```yaml
Type: String
Parameter Sets: Html, HtmlPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FontFamily
Optional font family for HTML to Word conversion before RTF serialization.

```yaml
Type: String
Parameter Sets: Html, HtmlPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Html
HTML markup to convert to RTF through the Word HTML converter.

```yaml
Type: String
Parameter Sets: Html
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HtmlPath
Path to an HTML file to convert to RTF.

```yaml
Type: String
Parameter Sets: HtmlPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OutputPath
Optional destination RTF path. When omitted, raw RTF text is returned.

```yaml
Type: String
Parameter Sets: WordPath, WordDocument, Html, HtmlPath, PdfPath
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
Parameter Sets: WordPath, WordDocument, Html, HtmlPath, PdfPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PdfPath
Path to a PDF file to convert to semantic RTF.

```yaml
Type: String
Parameter Sets: PdfPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -StylesheetContent
Inline CSS stylesheets to apply during HTML conversion.

```yaml
Type: String[]
Parameter Sets: Html, HtmlPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -StylesheetPath
Paths to CSS stylesheets to apply during HTML conversion.

```yaml
Type: String[]
Parameter Sets: Html, HtmlPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -WordDocument
Word document instance to convert to RTF.

```yaml
Type: WordDocument
Parameter Sets: WordDocument
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -WordPath
Path to a .docx file to convert to RTF.

```yaml
Type: String
Parameter Sets: WordPath
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

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `System.String
System.IO.FileInfo`

## RELATED LINKS

- None
