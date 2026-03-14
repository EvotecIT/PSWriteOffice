---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordWatermark
## SYNOPSIS
Adds a watermark to the current section or header.

## SYNTAX
### Text (Default)
```powershell
Add-OfficeWordWatermark [-Text] <string> [-HorizontalOffset <double>] [-VerticalOffset <double>] [-Scale <double>] [-PassThru] [<CommonParameters>]
```

### Image
```powershell
Add-OfficeWordWatermark [-ImagePath] <string> [-HorizontalOffset <double>] [-VerticalOffset <double>] [-Scale <double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a watermark to the current section or header.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Add-OfficeWordWatermark -Text 'CONFIDENTIAL'
```

Inserts a text watermark into the current section.

## PARAMETERS

### -HorizontalOffset
Horizontal offset for the watermark.

```yaml
Type: Nullable`1
Parameter Sets: Text, Image
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ImagePath
Path to an image watermark.

```yaml
Type: String
Parameter Sets: Image
Aliases: None
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the created watermark.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Image
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Scale
Scale factor for the watermark.

```yaml
Type: Double
Parameter Sets: Text, Image
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Text
Watermark text.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -VerticalOffset
Vertical offset for the watermark.

```yaml
Type: Nullable`1
Parameter Sets: Text, Image
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

- `System.Object`

## RELATED LINKS

- None

