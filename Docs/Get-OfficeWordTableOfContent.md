---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordTableOfContent
## SYNOPSIS
Gets the table of contents from a Word document.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeWordTableOfContent [-InputPath] <string> [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordTableOfContent -Document <WordDocument> [<CommonParameters>]
```

## DESCRIPTION
Gets the table of contents from a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficeWordTableOfContent -Path .\Report.docx
```

Returns the table of contents when one exists.

## PARAMETERS

### -Document
Word document to read.

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
Path to the .docx file.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordTableOfContent`

## RELATED LINKS

- None

