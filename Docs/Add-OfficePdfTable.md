---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfTable
## SYNOPSIS
Adds a table to a PDF document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfTable [-InputObject] <Object> [-Property <string[]>] [-Header <string[]>] [-View <OfficeTableView>] [-Align <PdfAlign>] [-TableStyle <string>] [-HeaderFill <string>] [-HeaderTextColor <string>] [-TextColor <string>] [-RowStripeFill <string>] [-BorderColor <string>] [-BorderWidth <double>] [-FontSize <double>] [-HeaderFontSize <double>] [-LineHeight <double>] [-CellPaddingX <double>] [-CellPaddingY <double>] [-SpacingBefore <double>] [-SpacingAfter <double>] [-Caption <string>] [-CaptionAlign <PdfAlign>] [-CaptionColor <string>] [-CaptionFontSize <double>] [-ColumnWidthPoints <double[]>] [-ColumnWidthWeights <double[]>] [-ColumnAlign <PdfColumnAlign[]>] [-AutoFitColumns] [-RightAlignNumeric] [-KeepTogether] [-KeepWithNext] [-NoBorder] [-NoHeaderFill] [-NoRowStripeFill] [-HeaderRowCount <int>] [-RepeatHeaderRowCount <int>] [-FooterRowCount <int>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfTable [-InputObject] <Object> -Document <PdfDocument> [-Property <string[]>] [-Header <string[]>] [-View <OfficeTableView>] [-Align <PdfAlign>] [-TableStyle <string>] [-HeaderFill <string>] [-HeaderTextColor <string>] [-TextColor <string>] [-RowStripeFill <string>] [-BorderColor <string>] [-BorderWidth <double>] [-FontSize <double>] [-HeaderFontSize <double>] [-LineHeight <double>] [-CellPaddingX <double>] [-CellPaddingY <double>] [-SpacingBefore <double>] [-SpacingAfter <double>] [-Caption <string>] [-CaptionAlign <PdfAlign>] [-CaptionColor <string>] [-CaptionFontSize <double>] [-ColumnWidthPoints <double[]>] [-ColumnWidthWeights <double[]>] [-ColumnAlign <PdfColumnAlign[]>] [-AutoFitColumns] [-RightAlignNumeric] [-KeepTogether] [-KeepWithNext] [-NoBorder] [-NoHeaderFill] [-NoRowStripeFill] [-HeaderRowCount <int>] [-RepeatHeaderRowCount <int>] [-FooterRowCount <int>] [-PassThru] [<CommonParameters>]
```

### PipelineDocument
```powershell
Add-OfficePdfTable [-InputObject] <Object> -Document <PdfDocument> [-Property <string[]>] [-Header <string[]>] [-View <OfficeTableView>] [-Align <PdfAlign>] [-TableStyle <string>] [-HeaderFill <string>] [-HeaderTextColor <string>] [-TextColor <string>] [-RowStripeFill <string>] [-BorderColor <string>] [-BorderWidth <double>] [-FontSize <double>] [-HeaderFontSize <double>] [-LineHeight <double>] [-CellPaddingX <double>] [-CellPaddingY <double>] [-SpacingBefore <double>] [-SpacingAfter <double>] [-Caption <string>] [-CaptionAlign <PdfAlign>] [-CaptionColor <string>] [-CaptionFontSize <double>] [-ColumnWidthPoints <double[]>] [-ColumnWidthWeights <double[]>] [-ColumnAlign <PdfColumnAlign[]>] [-AutoFitColumns] [-RightAlignNumeric] [-KeepTogether] [-KeepWithNext] [-NoBorder] [-NoHeaderFill] [-NoRowStripeFill] [-HeaderRowCount <int>] [-RepeatHeaderRowCount <int>] [-FooterRowCount <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a table to a PDF document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $services = @(
    [pscustomobject]@{ Name = 'Directory'; Status = 'Healthy'; Incidents = 0 }
    [pscustomobject]@{ Name = 'Mail'; Status = 'Watch'; Incidents = 2 }
)
New-OfficePdf -Path .\Examples\Documents\PdfTable.pdf {
    Add-OfficePdfHeading -Text 'Service status'
    Add-OfficePdfTable -InputObject $services -Property Name,Status,Incidents -Header 'Service','Status','Incidents'
}
```

Converts PowerShell objects into a table using selected properties and friendly headers.

## PARAMETERS

### -Align
Table alignment.

```yaml
Type: PdfAlign
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values: Left, Center, Right, Justify

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AutoFitColumns
Measure flexible columns from content.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -BorderColor
Border color in #RRGGBB format.

```yaml
Type: String
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -BorderWidth
Border width in PDF points.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Caption
Caption rendered above the table grid.

```yaml
Type: String
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CaptionAlign
Caption alignment.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CaptionColor
Caption color in #RRGGBB format.

```yaml
Type: String
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CaptionFontSize
Caption font size in PDF points.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CellPaddingX
Horizontal cell padding in PDF points.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CellPaddingY
Vertical cell padding in PDF points.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ColumnAlign
Per-column horizontal alignment.

```yaml
Type: PdfColumnAlign[]
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values: Left, Center, Right

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ColumnWidthPoints
Fixed column widths in PDF points.

```yaml
Type: Double[]
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ColumnWidthWeights
Relative column width weights.

```yaml
Type: Double[]
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
PDF document to update outside the DSL context.

```yaml
Type: PdfDocument
Parameter Sets: Document, PipelineDocument
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -FontSize
Body cell font size in PDF points.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FooterRowCount
Number of trailing rows rendered as footer rows.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Header
Header labels. Defaults to property names.

```yaml
Type: String[]
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HeaderFill
Header fill color in #RRGGBB format.

```yaml
Type: String
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HeaderFontSize
Header cell font size in PDF points.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HeaderRowCount
Number of leading rows rendered as header rows.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HeaderTextColor
Header text color in #RRGGBB format.

```yaml
Type: String
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -InputObject
Objects or row arrays to render as a table.

```yaml
Type: Object
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -KeepTogether
Keep the table together when possible.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -KeepWithNext
Keep the table with the next block when possible.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -LineHeight
Wrapped line height multiplier for table cells.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoBorder
Hide table borders.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoHeaderFill
Disable the header fill.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoRowStripeFill
Disable alternating row fill.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the updated document.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Property
Specific object properties to include.

```yaml
Type: String[]
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -RepeatHeaderRowCount
Number of leading header rows repeated on following pages.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -RightAlignNumeric
Right-align numeric-looking cell values.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -RowStripeFill
Alternating body row fill color in #RRGGBB format.

```yaml
Type: String
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SpacingAfter
Spacing after the table in PDF points.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SpacingBefore
Spacing before the table in PDF points.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -TableStyle
OfficeIMO table style preset or supported Word table style name.

```yaml
Type: String
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -TextColor
Body text color in #RRGGBB format.

```yaml
Type: String
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -View
Projection to apply before writing the table.

```yaml
Type: OfficeTableView
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values: Normal, Transpose

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Pdf.PdfDocument
System.Object`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
