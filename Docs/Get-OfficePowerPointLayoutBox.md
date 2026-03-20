---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePowerPointLayoutBox
## SYNOPSIS
Computes reusable layout boxes for a presentation.

## SYNTAX
### Content (Default)
```powershell
Get-OfficePowerPointLayoutBox [-Presentation <PowerPointPresentation>] [-MarginCm <double>] [<CommonParameters>]
```

### Columns
```powershell
Get-OfficePowerPointLayoutBox -ColumnCount <int> [-Presentation <PowerPointPresentation>] [-MarginCm <double>] [-GutterCm <double>] [<CommonParameters>]
```

### Rows
```powershell
Get-OfficePowerPointLayoutBox -RowCount <int> [-Presentation <PowerPointPresentation>] [-MarginCm <double>] [-GutterCm <double>] [<CommonParameters>]
```

## DESCRIPTION
Computes reusable layout boxes for a presentation.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficePowerPointLayoutBox -Presentation $ppt -MarginCm 1.5
```

Returns a single layout box representing the usable slide area.

### EXAMPLE 2
```powershell
PS>Get-OfficePowerPointLayoutBox -Presentation $ppt -ColumnCount 2 -MarginCm 1.5 -GutterCm 1.0
```

Returns one layout box per column.

## PARAMETERS

### -ColumnCount
Number of columns to generate.

```yaml
Type: Int32
Parameter Sets: Columns
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -GutterCm
Column or row gutter in centimeters.

```yaml
Type: Double
Parameter Sets: Columns, Rows
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MarginCm
Outer slide margin in centimeters.

```yaml
Type: Double
Parameter Sets: Content, Columns, Rows
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Presentation
Presentation to inspect (optional inside DSL).

```yaml
Type: PowerPointPresentation
Parameter Sets: Content, Columns, Rows
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -RowCount
Number of rows to generate.

```yaml
Type: Int32
Parameter Sets: Rows
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

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointLayoutBox`

## RELATED LINKS

- None

