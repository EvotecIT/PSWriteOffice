---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordImage
## SYNOPSIS
Inserts an image into the current paragraph.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordImage [-Path] <string> [-Width <double>] [-Height <double>] [-Wrap <WrapTextImage>] [-Description <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Inserts an image into the current paragraph.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Add-OfficeWordParagraph { Add-OfficeWordImage -Path .\logo.png -Width 96 -Height 32 }
```

Embeds logo.png at the specified size.

## PARAMETERS

### -Description
Optional description/alt text.

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

### -Height
Height in points.

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
Emit the created WordImage.

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
Path to the image file.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Width
Width in points.

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

### -Wrap
Wrap mode for the image.

```yaml
Type: WrapTextImage
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: InLineWithText, Square, Tight, Through, TopAndBottom, BehindText, InFrontOfText

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

