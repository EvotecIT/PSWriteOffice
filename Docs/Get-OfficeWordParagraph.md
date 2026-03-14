---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordParagraph
## SYNOPSIS
Gets paragraphs from a Word document or section.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeWordParagraph [-InputPath] <string> [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordParagraph -Document <WordDocument> [<CommonParameters>]
```

### Section
```powershell
Get-OfficeWordParagraph -Section <WordSection> [<CommonParameters>]
```

## DESCRIPTION
Gets paragraphs from a Word document or section.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficeWordParagraph -Path .\Report.docx
```

Returns all paragraphs in the document.

## PARAMETERS

### -Document
Document to inspect.

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

- `OfficeIMO.Word.WordDocument
OfficeIMO.Word.WordSection`

## OUTPUTS

- `OfficeIMO.Word.WordParagraph`

## RELATED LINKS

- None

