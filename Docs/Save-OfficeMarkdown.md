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
Save-OfficeMarkdown [-Document] <MarkdownDoc> [[-Path] <string>] [-PdfPath <string>] [-PassThru] [<CommonParameters>]
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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDoc
System.IO.FileInfo`

## RELATED LINKS

- None
