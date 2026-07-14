---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficePdfTableCell
## SYNOPSIS
Creates a reusable PDF table cell definition for explicit table rows.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficePdfTableCell [[-Text] <string>] [-Run <Object[]>] [-ColumnSpan <int>] [-RowSpan <int>] [-TextColor <string>] [-FillColor <string>] [-FontSize <double>] [-Bold] [-Italic] [-Underline] [-UnderlineStyle <string>] [-Strike] [-Align <PdfColumnAlign>] [-VerticalAlign <PdfCellVerticalAlign>] [<CommonParameters>]
```

## DESCRIPTION
Creates a reusable PDF table cell definition for explicit table rows.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $row = @(New-OfficePdfTableCell -Text 'Identity systems' -ColumnSpan 3 -FillColor '#DBEAFE' -TextColor '#1E3A8A' -Bold)
```

The returned cell can be passed to PdfTable inside explicit row arrays.

## PARAMETERS

### -Align
Horizontal cell alignment.

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

### -Bold
Render the cell text in bold.

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

### -ColumnSpan
Number of logical columns covered by the cell.

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

### -FillColor
Cell fill color. Named colors and hexadecimal colors are accepted.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: BackgroundColor, CellFill
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FontSize
Cell font size in PDF points.

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

### -Italic
Render the cell text in italics.

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

### -RowSpan
Number of logical rows covered by the cell.

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

### -Run
Rich text runs for the cell. Each run can be created with TextRun/PdfTextRun or provided as a hashtable/object.

```yaml
Type: Object[]
Parameter Sets: __AllParameterSets
Aliases: Runs
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Strike
Render the cell text with strikethrough.

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

### -Text
Cell text.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -TextColor
Cell text color. Named colors and hexadecimal colors are accepted.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: Color, FontColor
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Underline
Render the cell text with underline.

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

### -UnderlineStyle
Optional underline style name. PDF table rendering treats any supported value as underline.

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

### -VerticalAlign
Vertical cell alignment.

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

- `PSWriteOffice.Services.Table.OfficeTableCellSpec` — Describes a logical table cell that can be rendered by multiple Office table surfaces.

## RELATED LINKS

- None
