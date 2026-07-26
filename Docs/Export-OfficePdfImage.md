---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficePdfImage
## SYNOPSIS
Exports PDF pages through the shared PNG, JPEG, TIFF, SVG, or WebP image contract.

## SYNTAX
### __AllParameterSets
```powershell
Export-OfficePdfImage [-Path] <string> [-OutputPath] <string> [-PageRange <string>] [-Format <OfficeImageExportFormat>] [-Options <PdfImageExportOptions>] [-ReadOptions <PdfReadOptions>] [-Password <string>] [-IgnorePermissionRestrictions] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Exports PDF pages through the shared PNG, JPEG, TIFF, SVG, or WebP image contract.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Export-OfficePdfImage -Path .\Report.pdf -OutputPath .\Pages -PageRange '1-3,5'
```

Writes the selected pages and returns normalized image results with rendering diagnostics.

## PARAMETERS

### -Format
Output image format.

```yaml
Type: OfficeImageExportFormat
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Png, Svg, Jpeg, Tiff, Webp

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed extraction restrictions.

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

### -Options
Optional DPI, scale, thumbnail, encoding, diagnostics, and resource limits.

```yaml
Type: PdfImageExportOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OutputPath
Destination folder.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PageRange
Optional one-based ranges such as 1-3,5.

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

### -Password
Password used to authenticate an encrypted PDF.

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
Path to the PDF.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ReadOptions
Optional bounded PDF parsing settings.

```yaml
Type: PdfReadOptions
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

- `None`

## OUTPUTS

- `OfficeIMO.Drawing.OfficeImageExportResult`

## RELATED LINKS

- None
