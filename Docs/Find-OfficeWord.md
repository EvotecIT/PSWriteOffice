---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Find-OfficeWord
## SYNOPSIS
Finds text matches inside a Word document.

## SYNTAX
### PathText (Default)
```powershell
Find-OfficeWord [-InputPath] <string> [-Text] <string> [-CaseSensitive] [<CommonParameters>]
```

### PathRegex
```powershell
Find-OfficeWord [-InputPath] <string> [-Pattern] <string> [-CaseSensitive] [-AsResult] [<CommonParameters>]
```

### DocumentText
```powershell
Find-OfficeWord [-Text] <string> -Document <WordDocument> [-CaseSensitive] [<CommonParameters>]
```

### DocumentRegex
```powershell
Find-OfficeWord [-Pattern] <string> -Document <WordDocument> [-CaseSensitive] [-AsResult] [<CommonParameters>]
```

## DESCRIPTION
Finds text matches inside a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Find-OfficeWord -Path .\Report.docx -Text 'Quarter'
```

Returns paragraphs that contain the search text.

## PARAMETERS

### -AsResult
Emit the full WordFind result for regex searches.

```yaml
Type: SwitchParameter
Parameter Sets: PathRegex, DocumentRegex
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CaseSensitive
Use case-sensitive matching.

```yaml
Type: SwitchParameter
Parameter Sets: PathText, PathRegex, DocumentText, DocumentRegex
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Word document to search.

```yaml
Type: WordDocument
Parameter Sets: DocumentText, DocumentRegex
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -InputPath
Path to the .docx file.

```yaml
Type: String
Parameter Sets: PathText, PathRegex
Aliases: FilePath, Path
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Pattern
Regular expression pattern to find.

```yaml
Type: String
Parameter Sets: PathRegex, DocumentRegex
Aliases: None
Possible values: 

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Text
Text to find.

```yaml
Type: String
Parameter Sets: PathText, DocumentText
Aliases: None
Possible values: 

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordParagraph
OfficeIMO.Word.WordFind`

## RELATED LINKS

- None

