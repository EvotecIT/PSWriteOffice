---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficePdfRedacted
## SYNOPSIS
Applies rectangle-based PDF redactions and writes a new PDF.

## SYNTAX
### Rectangle (Default)
```powershell
ConvertTo-OfficePdfRedacted [-Path] <string> [-OutputPath] <string> -PageNumber <int> -X <double> -Y <double> -Width <double> -Height <double> [-Label <string>] [-FillColor <string>] [-OnlyPaintMatches] [-Password <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Area
```powershell
ConvertTo-OfficePdfRedacted [-Path] <string> [-OutputPath] <string> -Area <PdfRedactionArea[]> [-FillColor <string>] [-OnlyPaintMatches] [-Password <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Applies rectangle-based PDF redactions and writes a new PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ConvertTo-OfficePdfRedacted -Path .\Report.pdf -OutputPath .\Report-Redacted.pdf -PageNumber 1 -X 72 -Y 650 -Width 240 -Height 32
```

Removes matching text objects and annotations in the rectangle, then paints a redaction mark.

## PARAMETERS

### -Area
One or more pre-created OfficeIMO.Pdf redaction areas.

```yaml
Type: PdfRedactionArea[]
Parameter Sets: Area
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FillColor
Redaction fill color in #RRGGBB format. Defaults to black.

```yaml
Type: String
Parameter Sets: Rectangle, Area
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Height
Rectangle height in PDF points.

```yaml
Type: Double
Parameter Sets: Rectangle
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Label
Optional redaction area label.

```yaml
Type: String
Parameter Sets: Rectangle
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OnlyPaintMatches
Paint only areas that match text or annotations in the redaction plan.

```yaml
Type: SwitchParameter
Parameter Sets: Rectangle, Area
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
Parameter Sets: Rectangle, Area
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PageNumber
One-based page number for the redaction rectangle.

```yaml
Type: Int32
Parameter Sets: Rectangle
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Password
Password used to read a Standard password-encrypted PDF.

```yaml
Type: String
Parameter Sets: Rectangle, Area
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
Parameter Sets: Rectangle, Area
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Width
Rectangle width in PDF points.

```yaml
Type: Double
Parameter Sets: Rectangle
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -X
Left coordinate in PDF points.

```yaml
Type: Double
Parameter Sets: Rectangle
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Y
Bottom coordinate in PDF points.

```yaml
Type: Double
Parameter Sets: Rectangle
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
