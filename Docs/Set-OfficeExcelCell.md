---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelCell
## SYNOPSIS
Sets a cell value, formula, or number format within the current worksheet.

## SYNTAX
### Coordinates
```powershell
Set-OfficeExcelCell [-Row <Int32>] [-Column <Int32>] [-Value <Object>] [-Formula <string>] [-NumberFormat <string>] [-BackgroundColor <string>] [-GradientFrom <string>] [-GradientTo <string>] [-GradientDegree <double>] [<CommonParameters>]
```

### Address
```powershell
Set-OfficeExcelCell [-Address <string>] [-Value <Object>] [-Formula <string>] [-NumberFormat <string>] [-BackgroundColor <string>] [-GradientFrom <string>] [-GradientTo <string>] [-GradientDegree <double>] [<CommonParameters>]
```

## DESCRIPTION
Supports A1 addresses or row/column coordinates for DSL-style composition.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Set-OfficeExcelCell -Address 'A1' -Value 'Region'; Set-OfficeExcelCell -Row 1 -Column 2 -Value 'Revenue' }
```

Writes two headers in the first row.

## PARAMETERS

### -Address
A1-style cell address (e.g., A1, C5).

```yaml
Type: String
Parameter Sets: Address
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundColor
Solid background color as #RRGGBB or #AARRGGBB.

```yaml
Type: String
Parameter Sets: Coordinates, Address
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Column
1-based column index.

```yaml
Type: Int32
Parameter Sets: Coordinates
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Formula
Formula text (without leading =).

```yaml
Type: String
Parameter Sets: Coordinates, Address
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GradientDegree
Linear gradient angle in degrees.

```yaml
Type: Double
Parameter Sets: Coordinates, Address
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GradientFrom
Gradient start color as #RRGGBB or #AARRGGBB.

```yaml
Type: String
Parameter Sets: Coordinates, Address
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GradientTo
Gradient end color as #RRGGBB or #AARRGGBB.

```yaml
Type: String
Parameter Sets: Coordinates, Address
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NumberFormat
Number format code to apply.

```yaml
Type: String
Parameter Sets: Coordinates, Address
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Row
1-based row index.

```yaml
Type: Int32
Parameter Sets: Coordinates
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value
Cell value to assign.

```yaml
Type: Object
Parameter Sets: Coordinates, Address
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
