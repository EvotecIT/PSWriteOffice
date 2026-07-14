---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeTextRun
## SYNOPSIS
Creates a reusable rich text run specification for Word, Excel, PowerPoint, and PDF commands.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeTextRun [[-Text] <string>] [-Kind <string>] [-Bold] [-Italic] [-Underline] [-UnderlineStyle <string>] [-Strike] [-Color <string>] [-BackgroundColor <string>] [-FontSize <double>] [-FontName <string>] [-Baseline <string>] [-LinkUri <string>] [-LinkDestinationName <string>] [-LinkContents <string>] [-TabLeader <string>] [-TabAlignment <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates a reusable rich text run specification for Word, Excel, PowerPoint, and PDF commands.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeTextRun -BackgroundColor 'Value'
```


## PARAMETERS

### -BackgroundColor
Run background or highlight color. Named colors and hexadecimal colors are accepted.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: HighlightColor, FillColor
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Baseline
Target-specific baseline name, such as Superscript or Subscript.

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

### -Bold
Render the run in bold.

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

### -Color
Text color. Named colors and hexadecimal colors are accepted.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: TextColor, FontColor
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FontName
Font name, family, or target-specific font identifier.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: Font, Typeface, FontFamily
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

### -Italic
Render the run in italics.

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

### -Kind
Run kind such as Text, LineBreak, Tab, Superscript, or Subscript.

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

### -LinkContents
Optional link tooltip or annotation contents.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: Contents, Tooltip
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -LinkDestinationName
Named destination or bookmark target when supported by the target format.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: DestinationName, Bookmark, BookmarkName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -LinkUri
URI link target when supported by the target format.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: Uri, Url, Href
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Strike
Render the run with strikethrough.

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

### -TabAlignment
Tab alignment name.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: Alignment
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -TabLeader
PDF tab leader style name.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: Leader
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Text
Run text.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Underline
Render the run with underline.

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

### -UnderlineStyle
Optional underline style name when the target format supports it.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `PSWriteOffice.Services.Text.OfficeTextRunSpec` — PowerShell-friendly rich text run specification used by document adapters.

## RELATED LINKS

- None
