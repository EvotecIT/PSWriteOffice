---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfBookmark
## SYNOPSIS
Adds a named bookmark at the current generated PDF flow position.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfBookmark [-Name] <string> [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfBookmark [-Name] <string> -Document <PdfDocument> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a named bookmark at the current generated PDF flow position.

## EXAMPLES

### EXAMPLE 1
```powershell
Add-OfficePdfBookmark -Name 'Name'
```


### EXAMPLE 2
```powershell
Add-OfficePdfBookmark -Document 'Value'
```


## PARAMETERS

### -Document
PDF document to update outside the DSL context.

```yaml
Type: PdfDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Name
Bookmark name.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the updated document.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
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

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
