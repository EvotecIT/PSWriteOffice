---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficePdfLayoutOverlay
## SYNOPSIS
Exports PDF word, line, region, and reading-order diagnostics as PNG or SVG.

## SYNTAX
### __AllParameterSets
```powershell
Export-OfficePdfLayoutOverlay [-Path] <string> [-OutputPath] <string> [-Page <int>] [-Format <OfficeImageExportFormat>] [-Scale <double>] [-Options <PdfLayoutDebugOverlayOptions>] [-LayoutOptions <PdfTextLayoutOptions>] [-ReadOptions <PdfReadOptions>] [-Password <string>] [-IgnorePermissionRestrictions] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Exports PDF word, line, region, and reading-order diagnostics as PNG or SVG.

## EXAMPLES

### EXAMPLE 1
```powershell
Export-OfficePdfLayoutOverlay -Path 'C:\Path'
```


## PARAMETERS

### -Format
Output image format. Layout overlays support only PNG and SVG.

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

### -LayoutOptions
Optional text layout settings.

```yaml
Type: PdfTextLayoutOptions
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
Optional overlay elements, colors, and limits.

```yaml
Type: PdfLayoutDebugOverlayOptions
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
Destination PNG or SVG path.

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

### -Page
One-based page number.

```yaml
Type: Int32
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
Source PDF path.

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

### -Scale
Output scale.

```yaml
Type: Double
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
