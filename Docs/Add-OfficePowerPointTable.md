---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePowerPointTable
## SYNOPSIS
Adds a table to a PowerPoint slide.

## SYNTAX
### InputObject (Default)
```powershell
Add-OfficePowerPointTable [[-Slide] <PowerPointSlide>] [-InputObject] <Object> [-Header <string[]>] [-NoHeader] [-View <OfficeTableView>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-StyleId <string>] [<CommonParameters>]
```

### Size
```powershell
Add-OfficePowerPointTable [[-Slide] <PowerPointSlide>] -Rows <int> -Columns <int> [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-StyleId <string>] [<CommonParameters>]
```

## DESCRIPTION
Builds a table from data rows or creates a blank grid with a fixed size.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $rows = @([pscustomobject]@{ Item='Alpha'; Qty=2 }, [pscustomobject]@{ Item='Beta'; Qty=4 })
Add-OfficePowerPointTable -Slide $slide -InputObject $rows -X 60 -Y 140 -Width 420 -Height 200
```

Creates a table with headers and two data rows.

## PARAMETERS

### -Columns
Column count for an empty table.

```yaml
Type: Int32
Parameter Sets: Size
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Header
Optional header order to apply to the table.

```yaml
Type: String[]
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Height
Table height in points.

```yaml
Type: Double
Parameter Sets: InputObject, Size
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -InputObject
Source objects to convert into table rows.

```yaml
Type: Object
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoHeader
Skip writing header row.

```yaml
Type: SwitchParameter
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Rows
Row count for an empty table.

```yaml
Type: Int32
Parameter Sets: Size
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Slide
Target slide that will receive the table (optional inside DSL).

```yaml
Type: PowerPointSlide
Parameter Sets: InputObject, Size
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -StyleId
Optional table style ID (GUID string).

```yaml
Type: String
Parameter Sets: InputObject, Size
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
Parameter Sets: InputObject
Aliases: None
Possible values: Normal, Transpose

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Width
Table width in points.

```yaml
Type: Double
Parameter Sets: InputObject, Size
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -X
Left offset (in points) from the slide origin.

```yaml
Type: Double
Parameter Sets: InputObject, Size
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Y
Top offset (in points) from the slide origin.

```yaml
Type: Double
Parameter Sets: InputObject, Size
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

- `OfficeIMO.PowerPoint.PowerPointSlide`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None
