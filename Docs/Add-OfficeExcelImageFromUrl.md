---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelImageFromUrl
## SYNOPSIS
Adds an image from a URL anchored to a worksheet cell or range.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelImageFromUrl [-Url] <string> [-Row <int>] [-Column <int>] [-Address <string>] [-Range <string>] [-WidthPixels <int>] [-HeightPixels <int>] [-ScalePercent <double>] [-OffsetXPixels <int>] [-OffsetYPixels <int>] [-EndOffsetXPixels <int>] [-EndOffsetYPixels <int>] [-Name <string>] [-AltText <string>] [-Title <string>] [-Decorative] [-NoLockAspectRatio] [-Placement <ExcelImagePlacement>] [-RotationDegrees <double>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelImageFromUrl [-Url] <string> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-Row <int>] [-Column <int>] [-Address <string>] [-Range <string>] [-WidthPixels <int>] [-HeightPixels <int>] [-ScalePercent <double>] [-OffsetXPixels <int>] [-OffsetYPixels <int>] [-EndOffsetXPixels <int>] [-EndOffsetYPixels <int>] [-Name <string>] [-AltText <string>] [-Title <string>] [-Decorative] [-NoLockAspectRatio] [-Placement <ExcelImagePlacement>] [-RotationDegrees <double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds an image from a URL anchored to a worksheet cell or range.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Add-OfficeExcelImageFromUrl -Address 'B2' -Url 'https://example.org/logo.png' -ScalePercent 20 -Name Logo -AltText 'Company logo' }
```

Downloads the remote image, sizes it to 20 percent of its original dimensions, and anchors it to cell B2.

### EXAMPLE 2
```powershell
PS> ExcelSheet 'Data' { Add-OfficeExcelImageFromUrl -Range 'A1:C15' -Url 'https://example.org/logo.png' -Placement MoveAndSize }
```

Uses Excel's two-cell anchor so the image moves and resizes with the cells in A1:C15.

## PARAMETERS

### -Address
A1-style cell address (e.g., A1, C5).

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: Cell
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AltText
Optional alternative text description for accessibility.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Column
1-based column index.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Decorative
Marks the image as decorative by clearing alternative text metadata.

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

### -Document
Workbook to operate on outside the DSL context.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -EndOffsetXPixels
Horizontal offset in pixels for the range end marker when using Range.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -EndOffsetYPixels
Vertical offset in pixels for the range end marker when using Range.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HeightPixels
Image height in pixels.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: Height
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Name
Optional drawing name used by Excel's selection pane.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoLockAspectRatio
Do not lock the image aspect ratio in Excel.

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

### -OffsetXPixels
Horizontal offset in pixels from the cell origin.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OffsetYPixels
Vertical offset in pixels from the cell origin.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the worksheet after inserting the image.

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

### -Placement
How a range-anchored image behaves when cells move or resize.

```yaml
Type: ExcelImagePlacement
Parameter Sets: Context, Document
Aliases: None
Possible values: MoveAndSize, MoveOnly, FreeFloating

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Range
A1-style range (for example, A1:C15) for a two-cell anchor that can move and resize with cells.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -RotationDegrees
Clockwise image rotation in degrees.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Row
1-based row index.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ScalePercent
Percentage of the original image size. Cannot be combined with WidthPixels or HeightPixels.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Sheet
Worksheet name when using Document.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SheetIndex
Worksheet index (0-based) when using Document.

```yaml
Type: Nullable`1
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Title
Optional alternative text title.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Url
Image URL to download.

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

### -WidthPixels
Image width in pixels.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: Width
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

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None
