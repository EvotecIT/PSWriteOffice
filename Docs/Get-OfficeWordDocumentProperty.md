---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordDocumentProperty
## SYNOPSIS
Gets built-in and custom document properties from a Word document.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeWordDocumentProperty [-InputPath] <string> [-Name <string[]>] [-BuiltIn] [-Custom] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordDocumentProperty -Document <WordDocument> [-Name <string[]>] [-BuiltIn] [-Custom] [<CommonParameters>]
```

## DESCRIPTION
Gets built-in and custom document properties from a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficeWordDocumentProperty -Path .\Report.docx
```

Returns built-in and custom Word document properties.

## PARAMETERS

### -BuiltIn
Only return built-in document properties.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Custom
Only return custom document properties.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

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

### -Name
Property name filter (wildcards supported).

```yaml
Type: String[]
Parameter Sets: Path, Document
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

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `PSWriteOffice.Models.Word.WordDocumentPropertyInfo`

## RELATED LINKS

- None

