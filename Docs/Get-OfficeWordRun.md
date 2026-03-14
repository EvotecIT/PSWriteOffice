---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordRun
## SYNOPSIS
Gets runs from Word paragraphs.

## SYNTAX
### Paragraph (Default)
```powershell
Get-OfficeWordRun -Paragraph <WordParagraph> [<CommonParameters>]
```

### Section
```powershell
Get-OfficeWordRun -Section <WordSection> [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordRun -Document <WordDocument> [<CommonParameters>]
```

### Path
```powershell
Get-OfficeWordRun [-InputPath] <string> [<CommonParameters>]
```

## DESCRIPTION
Gets runs from Word paragraphs.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficeWordParagraph -Path .\Report.docx | Get-OfficeWordRun
```

Returns each run as a WordParagraph instance.

## PARAMETERS

### -Document
Document to enumerate.

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

### -InputPath
Path to the document.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath, Path
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Paragraph
Paragraph to enumerate.

```yaml
Type: WordParagraph
Parameter Sets: Paragraph
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Section
Section to enumerate.

```yaml
Type: WordSection
Parameter Sets: Section
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordParagraph
OfficeIMO.Word.WordSection
OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordParagraph`

## RELATED LINKS

- None

