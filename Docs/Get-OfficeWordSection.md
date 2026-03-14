---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordSection
## SYNOPSIS
Gets sections from a Word document.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeWordSection [-InputPath] <string> [-Index <int[]>] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordSection -Document <WordDocument> [-Index <int[]>] [<CommonParameters>]
```

## DESCRIPTION
Gets sections from a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficeWordSection -Path .\Report.docx
```

Returns the sections contained in the document.

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

### -Index
Optional 0-based section index filter.

```yaml
Type: Int32[]
Parameter Sets: Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordSection`

## RELATED LINKS

- None

