---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Join-OfficePdf
## SYNOPSIS
Joins multiple PDF files into a single PDF.

## SYNTAX
### __AllParameterSets
```powershell
Join-OfficePdf [-Path] <string[]> [-OutputPath] <string> [-PassThru] [-FlattenVisualAnnotations] [-PageSize <string>] [-Width <double>] [-Height <double>] [-Landscape] [-ResizeMode <PdfPageResizeMode>] [-ResizeMargin <double>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Joins multiple PDF files into a single PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $cover = '.\Examples\Documents\Cover.pdf'
$report = '.\Examples\Documents\Report.pdf'
Join-OfficePdf -Path $cover, $report -OutputPath .\Examples\Documents\Combined.pdf -PassThru
Get-OfficePdfInfo -Path .\Examples\Documents\Combined.pdf | Select-Object PageCount
```

Writes a single PDF containing the input documents in the requested order, then checks the result.

## PARAMETERS

### -FlattenVisualAnnotations
Flatten visual annotation appearances before merging.

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

### -Height
Custom output page height in points when -PageSize Custom is used.

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

### -Landscape
Use the landscape orientation of the selected output page size.

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

### -OutputPath
Output PDF path.

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

### -PageSize
Resize each merged page to a known OfficeIMO page size such as A4, Letter, or Custom.

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

### -PassThru
Emit the saved file.

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
Input PDF paths in output order.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
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

### -Width
Custom output page width in points when -PageSize Custom is used.

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
