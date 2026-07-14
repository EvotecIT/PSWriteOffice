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
Add-OfficeWordText [-Text] <string[]> [-Bold] [-Italic] [-Underline <UnderlineValues>] [-Color <string>] [-Strike] [-FontSize <int>] [-FontName <string>] [<CommonParameters>]
```

### Run
```powershell
Add-OfficeWordText -Run <Object[]> [-Bold] [-Italic] [-Underline <UnderlineValues>] [-Color <string>] [-Strike] [-FontSize <int>] [-FontName <string>] [<CommonParameters>]
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
Accept wildcard characters: True
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
Accept wildcard characters: True
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
Accept wildcard characters: True
```

### -FontSize
Font size in points.

```yaml
Type: Nullable`1
Parameter Sets: Text, Run
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
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
Accept wildcard characters: True
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
Accept wildcard characters: True
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
Accept wildcard characters: True
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
Accept wildcard characters: True
```

### -Underline
Optional underline style.

```yaml
Type: Nullable`1
Parameter Sets: Text, Run
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
