---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePdfPage
## SYNOPSIS
Sets page-level PDF properties and writes a new PDF.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficePdfPage -Path <string> -OutputPath <string> [-PageRange <string>] [-Rotation <int>] [-BoxName <string>] [-Left <double>] [-Bottom <double>] [-Right <double>] [-Top <double>] [-PageSize <string>] [-Width <double>] [-Height <double>] [-Landscape] [-ResizeMode <PdfPageResizeMode>] [-ResizeMargin <double>] [-Password <string>] [-IgnorePermissionRestrictions] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Sets page-level PDF properties and writes a new PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $proof = @(
    Set-OfficePdfPage -Path .\Examples\Documents\Scanned.pdf -PageRange '2,4' -Rotation 90 -OutputPath .\Examples\Documents\Scanned-Rotated.pdf
    Get-OfficePdfInfo -Path .\Examples\Documents\Scanned-Rotated.pdf | Select-Object PageCount
)
$proof
```

Rotates selected pages and writes a new PDF.

## PARAMETERS

### -Bottom
Bottom coordinate for the page boundary box.

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

### -BoxName
Page boundary box to set. Supported values are MediaBox, CropBox, BleedBox, TrimBox, and ArtBox.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: MediaBox, CropBox, BleedBox, TrimBox, ArtBox

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Height
Custom page height in points when -PageSize Custom is used.

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

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed page-modification restrictions.

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

### -Landscape
Use the landscape orientation of the selected page size.

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

### -Left
Left coordinate for the page boundary box.

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
Page ranges such as 1-3,5. Omit to affect all pages.

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

### -PageSize
Resize selected pages to a known OfficeIMO page size such as A4, Letter, or Custom.

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
Input PDF path.

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

### -ResizeMargin
Margin, in points, reserved around resized page content.

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

### -ResizeMode
How source page content is fitted into the resized output page.

```yaml
Type: PdfPageResizeMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Fit, Fill, Stretch

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Right
Right coordinate for the page boundary box.

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

### -Rotation
Rotation in degrees. Supported values are 0, 90, 180, and 270.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 0, 90, 180, 270

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Top
Top coordinate for the page boundary box.

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

### -Width
Custom page width in points when -PageSize Custom is used.

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
