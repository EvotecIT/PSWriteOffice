---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointLayoutPlaceholderTextStyle
## SYNOPSIS
Sets layout placeholder text style and bullet/numbering settings.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficePowerPointLayoutPlaceholderTextStyle -Layout <int> -PlaceholderType <string> [-Presentation <PowerPointPresentation>] [-Master <int>] [-Index <uint>] [-Style <string>] [-FontSize <int>] [-FontName <string>] [-Color <string>] [-Bold <bool>] [-Italic <bool>] [-Underline <bool>] [-HighlightColor <string>] [-Level <int>] [-BulletChar <string>] [-Numbering <string>] [-CreateIfMissing] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets layout placeholder text style and bullet/numbering settings.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Set-OfficePowerPointLayoutPlaceholderTextStyle -Presentation $ppt -Master 0 -Layout 1 -PlaceholderType Title -Style Title
```

Applies the Title preset to the layout placeholder.

### EXAMPLE 2
```powershell
PS>New-OfficePowerPoint -Path .\deck.pptx {
$layout = Get-OfficePowerPointLayout | Select-Object -First 1
Set-OfficePowerPointLayoutPlaceholderTextStyle -Master $layout.MasterIndex -Layout $layout.LayoutIndex -PlaceholderType Title -Style Title -FontSize 36 -Bold $true
}
```

Uses the DSL context to resolve the presentation.

## PARAMETERS

### -Bold
Apply bold formatting.

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

### -BulletChar
Optional bullet character (ignored when -Numbering is supplied).

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

### -Color
Text color in hex (e.g. 1F4E79).

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

### -CreateIfMissing
Create the placeholder if it is missing.

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

### -FontName
Font name (Latin).

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

### -FontSize
Font size in points.

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

### -HighlightColor
Highlight color in hex (e.g. FFF59D).

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

### -Index
Optional placeholder index.

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
Apply italic formatting.

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

### -Layout
Layout index within the master.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Level
Paragraph level (0-8) to set before applying style.

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

### -Master
Slide master index.

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

### -Numbering
Optional numbering scheme name (e.g. ArabicPeriod, RomanUpper).

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

### -PassThru
Emit the placeholder textbox after update.

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

### -PlaceholderType
Placeholder type to target.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: Type
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Presentation
Presentation to update (optional inside DSL).

```yaml
Type: PowerPointPresentation
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Style
Named style preset (Title, Subtitle, Body, Caption, Emphasis).

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

### -Underline
Apply underline formatting.

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

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointTextBox`

## RELATED LINKS

- None

