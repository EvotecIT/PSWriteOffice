---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePdfImage
## SYNOPSIS
Gets or extracts image resources from a PDF.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePdfImage [-Path] <string> [-PageRange <string>] [-OutputDirectory <string>] [-BaseName <string>] [<CommonParameters>]
```

## DESCRIPTION
Gets or extracts image resources from a PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePdfImage -Path .\Examples\Documents\Report.pdf -PageRange '1-2'
            Get-OfficePdfImage -Path .\Examples\Documents\Report.pdf -OutputDirectory .\Examples\Documents\PdfImages -BaseName 'report-image'
```

Returns image metadata or writes extracted images to disk.

## PARAMETERS

### -BaseName
Base file name used when extracting images to disk.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OutputDirectory
Optional directory where images should be written.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PageRange
Optional page ranges such as 1-3,5.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

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

- `OfficeIMO.Pdf.PdfExtractedImage
System.IO.FileInfo`

## RELATED LINKS

- None
