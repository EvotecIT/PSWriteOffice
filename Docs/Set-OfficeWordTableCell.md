---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeWordTableCell
## SYNOPSIS
Updates OfficeIMO Word table-cell content, layout, and merge settings.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeWordTableCell [-Cell] <WordTableCell> [-Text <string>] [-ShadingFillColor <string>] [-ShadingPattern <string>] [-Width <int>] [-WidthType <string>] [-TextDirection <TextDirectionValues>] [-WrapText <bool>] [-FitText <bool>] [-MergeRight <int>] [-MergeDown <int>] [-SplitHorizontal <int>] [-SplitVertical <int>] [-CopyParagraphs] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Updates OfficeIMO Word table-cell content, layout, and merge settings.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $doc = Get-OfficeWord -Path .\Handover.docx
$table = Find-OfficeWordTable -Document $doc -Text 'Risk marker' | Select-Object -First 1
$table |
    Get-OfficeWordTableCell -Row 2 -Column 2 |
    Set-OfficeWordTableCell -Text 'Investigating' -ShadingFillColor '#fff2cc' -ShadingPattern Clear
$doc | Close-OfficeWord -Save
```

Finds an existing table by text, replaces a target cell value, applies shading, and saves the document.

### EXAMPLE 2
```powershell
PS> $doc = Get-OfficeWord -Path .\Report.docx
$table = $doc | Get-OfficeWordTable | Select-Object -First 1
$table |
    Get-OfficeWordTableCell -Column 2 |
    Set-OfficeWordTableCell -ShadingFillColor '#fff1f0' -ShadingPattern Clear -Width 2400 -WidthType Dxa
$doc | Save-OfficeWord -Path .\Report-StatusCells.docx
```

Reads cells from an OfficeIMO table object, applies cell shading and width, and saves the updated document.

### EXAMPLE 3
```powershell
PS> $doc = Get-OfficeWord -Path .\Report.docx
$table = $doc | Get-OfficeWordTable | Select-Object -First 1
$table |
    Get-OfficeWordTableCell -Row 0 -Column 0 |
    Set-OfficeWordTableCell -MergeRight 2 -CopyParagraphs
$doc | Save-OfficeWord -Path .\Report-MergedHeader.docx
```

Uses the OfficeIMO merge operation exposed by the thin table-cell wrapper.

## PARAMETERS

### -Cell
Table cell to update.

```yaml
Type: WordTableCell
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -CopyParagraphs
Copy paragraphs while merging cells.

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

### -FitText
Whether text should fit within the cell.

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

### -MergeDown
Number of cells to merge downward.

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

### -MergeRight
Number of cells to merge to the right.

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

### -PassThru
Emit the updated table cell.

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

### -ShadingFillColor
Cell shading fill color as #RRGGBB.

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

### -ShadingPattern
Cell shading pattern.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Nil, Clear, Solid, HorizontalStripe, VerticalStripe, ReverseDiagonalStripe, DiagonalStripe, HorizontalCross, DiagonalCross, ThinHorizontalStripe, ThinVerticalStripe, ThinReverseDiagonalStripe, ThinDiagonalStripe, ThinHorizontalCross, ThinDiagonalCross, Percent5, Percent10, Percent12, Percent15, Percent20, Percent25, Percent30, Percent35, Percent37, Percent40, Percent45, Percent50, Percent55, Percent60, Percent62, Percent65, Percent70, Percent75, Percent80, Percent85, Percent87, Percent90, Percent95

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SplitHorizontal
Number of columns to split the cell into.

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

### -SplitVertical
Number of rows to split the cell into.

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

### -Text
Replace the visible cell text.

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

### -TextDirection
Cell text direction.

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
Cell width value.

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

### -WidthType
Cell width unit type.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Nil, Pct, Dxa, Auto

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -WrapText
Whether text wraps in the cell.

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

- `OfficeIMO.Word.WordTableCell`

## OUTPUTS

- `OfficeIMO.Word.WordTableCell`

## RELATED LINKS

- None
