---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficePdfSearchable
## SYNOPSIS
Creates a searchable PDF by adding invisible text from a discovered local OCR runtime.

## SYNTAX
### __AllParameterSets
```powershell
ConvertTo-OfficePdfSearchable [-Path] <string> [-OutputPath] <string> [-Force] [-PassThru] [-RenderDpi <Double>] [-MinimumConfidence <Double>] [-Options <OfficeOcrOptions>] [-Language <string>] [-TesseractPath <string>] [-TessdataDirectory <string>] [-NoLanguageDownload] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Creates a searchable PDF by adding invisible text from a discovered local OCR runtime.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ConvertTo-OfficePdfSearchable -Path .\Scan.pdf -OutputPath .\Scan-Searchable.pdf
```

Preserves visible page content and writes geometry-aligned invisible English text.

### EXAMPLE 2
```powershell
PS> ConvertTo-OfficePdfSearchable -Path .\Scan.pdf -OutputPath .\Searchable.pdf -Language eng+pol -PassThru
```

Returns recognition, filtering, page, provider, and model evidence instead of the output file.

## PARAMETERS

### -Force
Overwrite an existing destination file.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Language
Tesseract language expression, such as eng or eng+pol. The default is eng.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MinimumConfidence
Minimum normalized confidence accepted for searchable text.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoLanguageDownload
Do not download checksum-pinned curated language data when a requested language is missing.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Advanced OfficeIMO OCR options. Convenience parameters override matching values.

```yaml
Type: OfficeOcrOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Destination PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the complete searchable-PDF OCR result instead of the output file.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Source PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -RenderDpi
PDF page render resolution used for recognition.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TessdataDirectory
Explicit directory containing Tesseract trained-data files.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TesseractPath
Explicit Tesseract executable path. By default OfficeIMO securely discovers an installed runtime.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `System.IO.FileInfo`
- `OfficeIMO.Pdf.PdfSearchableOcrResult`

## RELATED LINKS

- None
