---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfPageOverlay
## SYNOPSIS
Overlays or underlays one source PDF page on selected pages of another PDF.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficePdfPageOverlay -Path <string> -SourcePath <string> -OutputPath <string> [-SourcePageNumber <int>] [-PageRange <string>] [-Fit <PdfPageOverlayFit>] [-HorizontalAlign <PdfAlign>] [-VerticalAlign <PdfVerticalAlign>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-Opacity <double>] [-Underlay] [-Password <string>] [-IgnorePermissionRestrictions] [-SourcePassword <string>] [-IgnoreSourcePermissionRestrictions] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Overlays or underlays one source PDF page on selected pages of another PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficePdfPageOverlay -Path .\Report.pdf -SourcePath .\Letterhead.pdf `
                -SourcePageNumber 1 -Underlay -Opacity 0.9 -OutputPath .\BrandedReport.pdf
```

Imports the first source page once and places it behind each target page.

## PARAMETERS

### -Fit
How the source page fits the target page or explicit rectangle.

```yaml
Type: PdfPageOverlayFit
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: None, Contain, Cover, Stretch

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Height
Optional target rectangle height in PDF points.

```yaml
Type: Nullable`1
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HorizontalAlign
Horizontal placement inside the target page or rectangle.

```yaml
Type: PdfAlign
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Left, Center, Right, Justify

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -IgnorePermissionRestrictions
After target authentication, explicitly ignore owner-imposed target restrictions.

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

### -IgnoreSourcePermissionRestrictions
After source authentication, explicitly ignore owner-imposed source restrictions.

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

### -Opacity
Opacity of the imported source page.

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

### -OutputPath
Output PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PageRange
Target page selector such as 1-3,odd,last. Omit to apply to every target page.

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
Password used to authenticate the target PDF.

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
Target PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SourcePageNumber
One-based page number imported from the source PDF.

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

### -SourcePassword
Password used to authenticate the imported source PDF.

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

### -SourcePath
PDF containing the page to import.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Underlay
Place the imported page behind existing target-page content.

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

### -VerticalAlign
Vertical placement inside the target page or rectangle.

```yaml
Type: PdfVerticalAlign
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Top, Middle, Bottom

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Width
Optional target rectangle width in PDF points.

```yaml
Type: Nullable`1
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -X
Optional target rectangle X coordinate in PDF points.

```yaml
Type: Nullable`1
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Y
Optional target rectangle Y coordinate in PDF points.

```yaml
Type: Nullable`1
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

- `System.IO.FileInfo`

## RELATED LINKS

- None
