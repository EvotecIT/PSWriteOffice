---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordText
## SYNOPSIS
Adds inline text to the current paragraph.

## SYNTAX
### Text (Default)
```powershell
Add-OfficeWordText [-Text] <string[]> [-Bold] [-Italic] [-Underline <WordUnderlineStyle>] [-Color <string>] [-Strike] [-FontSize <Int32>] [-FontName <string>] [<CommonParameters>]
```

### Run
```powershell
Add-OfficeWordText -Run <Object[]> [-Bold] [-Italic] [-Underline <WordUnderlineStyle>] [-Color <string>] [-Strike] [-FontSize <Int32>] [-FontName <string>] [<CommonParameters>]
```

## DESCRIPTION
Supports bold/italic/underline and color tweaks for quick DSL composition.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeWordParagraph { Add-OfficeWordText -Text 'Important: ' -Bold }
```

Writes “Important:” with bold formatting.

## PARAMETERS

### -Bold
Apply bold formatting.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Run
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Run color (#RRGGBB).

```yaml
Type: String
Parameter Sets: Text, Run
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontName
Font name or family.

```yaml
Type: String
Parameter Sets: Text, Run
Aliases: Font, FontFamily, Typeface
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontSize
Font size in points.

```yaml
Type: Int32
Parameter Sets: Text, Run
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
Apply italic formatting.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Run
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Run
Rich text runs. Each run can be created with TextRun/WordTextRun or provided as a hashtable/object.

```yaml
Type: Object[]
Parameter Sets: Run
Aliases: Runs
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strike
Render text with strikethrough.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Run
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Text segments to append.

```yaml
Type: String[]
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Optional underline style.

```yaml
Type: WordUnderlineStyle
Parameter Sets: Text, Run
Aliases: None
Possible values: Single, Words, Double, Thick, Dotted, DottedHeavy, Dash, DashedHeavy, DashLong, DashLongHeavy, DotDash, DashDotHeavy, DotDotDash, DashDotDotHeavy, Wave, WavyHeavy, WavyDouble, None

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
