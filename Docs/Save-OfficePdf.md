---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Save-OfficePdf
## SYNOPSIS
Saves an OfficeIMO.Pdf document.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficePdf [-Document] <PdfDocument> [-Path] <string> [-Show] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Use this command when a PDF is built in memory and saved later, or when a pipeline should continue with the saved file.
The document is saved through the normal OfficeIMO.Pdf save path.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $pdf = New-OfficePdf { PdfHeading 'Queued report'; PdfParagraph 'Generated in memory.' }
            $pdf | Save-OfficePdf -Path .\QueuedReport.pdf
```

Creates a PDF document object first, then saves it to disk.

## PARAMETERS

### -Document
PDF document to save.

```yaml
Type: PdfDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -PassThru
Emit the document instead of the saved file.

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

### -Path
Destination PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Show
Open the PDF after saving.

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

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument
System.IO.FileInfo`

## RELATED LINKS

- None
