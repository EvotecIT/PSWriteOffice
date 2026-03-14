---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Remove-OfficeWordTableOfContent
## SYNOPSIS
Removes the table of contents from a Word document.

## SYNTAX
### __AllParameterSets
```powershell
Remove-OfficeWordTableOfContent [-Document <WordDocument>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Removes the table of contents from a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Remove-OfficeWordTableOfContent
```

Deletes the table of contents if one exists.

## PARAMETERS

### -Document
Document to modify when provided explicitly.

```yaml
Type: WordDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -PassThru
Emit the document after removal.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None

