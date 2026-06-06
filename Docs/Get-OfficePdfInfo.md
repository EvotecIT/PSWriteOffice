---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePdfInfo
## SYNOPSIS
Gets PDF metadata, page information, forms, links, and structural flags.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePdfInfo [-Path] <string> [<CommonParameters>]
```

## DESCRIPTION
The returned OfficeIMO.Pdf inspection object is useful for validating generated artifacts, migration scripts,
and existing PDFs before follow-up operations such as splitting, stamping, or metadata updates.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $info = Get-OfficePdfInfo -Path .\Report.pdf
            $info.PageCount
            $info.LinkUris
```

Reads page count and link information from the PDF.

## PARAMETERS

### -Path
PDF file path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocumentInfo`

## RELATED LINKS

- None
