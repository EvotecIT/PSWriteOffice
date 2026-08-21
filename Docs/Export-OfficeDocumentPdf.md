---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficeDocumentPdf
## SYNOPSIS
Exports a Word, Excel, PowerPoint, Markdown, or RTF document to PDF.

## SYNTAX
### Document (Default)
```powershell
Export-OfficeDocumentPdf [-Document] <Object> [-Path] <string> [-WordOptions <WordPdfSaveOptions>] [-ExcelOptions <ExcelPdfSaveOptions>] [-PowerPointOptions <PowerPointPdfSaveOptions>] [-MarkdownOptions <MarkdownPdfSaveOptions>] [-RtfOptions <RtfPdfSaveOptions>] [-PdfWarningVariable <string>] [-PdfConversionReportVariable <string>] [-Open] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Export-OfficeDocumentPdf [-InputPath] <string> [-Path] <string> [-Password <string>] [-WordOptions <WordPdfSaveOptions>] [-ExcelOptions <ExcelPdfSaveOptions>] [-PowerPointOptions <PowerPointPdfSaveOptions>] [-MarkdownOptions <MarkdownPdfSaveOptions>] [-RtfOptions <RtfPdfSaveOptions>] [-PdfWarningVariable <string>] [-PdfConversionReportVariable <string>] [-Open] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Accepts either a live OfficeIMO document from the pipeline or a supported source file.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $document | Export-OfficeDocumentPdf -Path .\Report.pdf
```


### EXAMPLE 2
```powershell
PS> Export-OfficeDocumentPdf -InputPath .\Report.docx -Path .\Report.pdf -PassThru
```


### EXAMPLE 3
```powershell
PS> $options = New-OfficeMarkdownPdfOptions -Title 'Service report' -IncludeLocalImages -BaseDirectory .\Assets
Export-OfficeDocumentPdf -InputPath .\Report.md -Path .\Report.pdf -MarkdownOptions $options
```

The New-Office*PdfOptions commands build every format-specific options object; no hashtable or .NET constructor is required.

## PARAMETERS

### -Document
Live Word, Excel, PowerPoint, Markdown, or RTF document to export. Saved FileInfo and path strings from the pipeline are opened automatically.

```yaml
Type: Object
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ExcelOptions
Excel-specific PDF options.

```yaml
Type: ExcelPdfSaveOptions
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputPath
Source .docx, .xlsx, .pptx, .md, .markdown, or .rtf file.

```yaml
Type: String
Parameter Sets: Path
Aliases: SourcePath, FullName
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -MarkdownOptions
Markdown-specific PDF options.

```yaml
Type: MarkdownPdfSaveOptions
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Open
Open the PDF after exporting it.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path
Aliases: Show
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the saved PDF file.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to open an encrypted Word, Excel, or PowerPoint source file.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Destination PDF path.

```yaml
Type: String
Parameter Sets: Document, Path
Aliases: OutputPath, FilePath
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfConversionReportVariable
Variable name that receives the structured PDF conversion report.

```yaml
Type: String
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfWarningVariable
Variable name that receives structured PDF conversion warnings.

```yaml
Type: String
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PowerPointOptions
PowerPoint-specific PDF options.

```yaml
Type: PowerPointPdfSaveOptions
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RtfOptions
RTF-specific PDF options.

```yaml
Type: RtfPdfSaveOptions
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WordOptions
Word-specific PDF options.

```yaml
Type: WordPdfSaveOptions
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.Object`
- `System.String`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
