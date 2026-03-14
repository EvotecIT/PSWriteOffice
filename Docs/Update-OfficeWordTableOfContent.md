---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Update-OfficeWordTableOfContent
## SYNOPSIS
Updates the table of contents in a Word document.

## SYNTAX
### TableOfContent
```powershell
Update-OfficeWordTableOfContent [-TableOfContent <WordTableOfContent>] [-Regenerate] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Update-OfficeWordTableOfContent [-Document <WordDocument>] [-Regenerate] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Updates the table of contents in a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Update-OfficeWordTableOfContent
```

Marks TOC fields as dirty and updates the document settings.

## PARAMETERS

### -Document
Document to update when provided explicitly.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -PassThru
Emit the updated table of contents.

```yaml
Type: SwitchParameter
Parameter Sets: TableOfContent, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Regenerate
Rebuild the table of contents before updating.

```yaml
Type: SwitchParameter
Parameter Sets: TableOfContent, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -TableOfContent
Table of contents to update.

```yaml
Type: WordTableOfContent
Parameter Sets: TableOfContent
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordTableOfContent
OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordTableOfContent`

## RELATED LINKS

- None

