---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordTextBox
## SYNOPSIS
Adds an OfficeIMO Word text box to the current Word DSL location.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordTextBox [-Text] <string> [-WrapText <WrapTextImage>] [-WidthCentimeters <double>] [-HeightCentimeters <double>] [-HorizontalOffsetCentimeters <double>] [-VerticalOffsetCentimeters <double>] [-HorizontalAlignment <WordHorizontalAlignmentValues>] [-HorizontalPositionRelativeFrom <HorizontalRelativePositionValues>] [-VerticalPositionRelativeFrom <VerticalRelativePositionValues>] [-AutoFit <WordTextBoxAutoFitType>] [-AutoFitToTextSize] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds an OfficeIMO Word text box to the current Word DSL location.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\Report.docx {
    WordTextBox -Text 'Executive summary' -WidthCentimeters 7 -HeightCentimeters 2 -HorizontalOffsetCentimeters 1.5 -VerticalOffsetCentimeters 1 -AutoFitToTextSize
}
```

Creates a native OfficeIMO Word text box and applies sizing/positioning through OfficeIMO's text-box API.

## PARAMETERS

### -AutoFit
Explicit OfficeIMO text-box autofit mode.

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

### -AutoFitToTextSize
Resize the text box to fit its text.

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

### -HeightCentimeters
Height in centimeters.

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

### -HorizontalAlignment
Horizontal alignment for anchored text boxes.

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

### -HorizontalOffsetCentimeters
Horizontal offset in centimeters.

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

### -HorizontalPositionRelativeFrom
Horizontal relative position anchor.

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
Emit the created text box.

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
Text to place inside the text box.

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

### -VerticalOffsetCentimeters
Vertical offset in centimeters.

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

### -VerticalPositionRelativeFrom
Vertical relative position anchor.

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

### -WidthCentimeters
Width in centimeters.

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

### -WrapText
Word text wrapping mode.

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

- `OfficeIMO.Word.WordTextBox`

## RELATED LINKS

- None
