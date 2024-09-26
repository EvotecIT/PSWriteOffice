---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version:
schema: 2.0.0
---

# New-OfficeWordText

## SYNOPSIS
{{ Fill in the Synopsis }}

## SYNTAX

### Document (Default)
```
New-OfficeWordText -Document <WordDocument> [-Text <String[]>] [-Bold <Nullable`1[]>] [-Italic <Nullable`1[]>]
 [-Underline <Nullable`1[]>] [-Color <String[]>] [-Alignment <JustificationValues>]
 [-Style <WordParagraphStyles>] [-ReturnObject] [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

### Paragraph
```
New-OfficeWordText [-Document <WordDocument>] -Paragraph <WordParagraph> [-Text <String[]>]
 [-Bold <Nullable`1[]>] [-Italic <Nullable`1[]>] [-Underline <Nullable`1[]>] [-Color <String[]>]
 [-Alignment <JustificationValues>] [-Style <WordParagraphStyles>] [-ReturnObject]
 [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -Alignment
{{ Fill Alignment Description }}

```yaml
Type: JustificationValues
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold
{{ Fill Bold Description }}

```yaml
Type: Nullable`1[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
{{ Fill Color Description }}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
{{ Fill Document Description }}

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

```yaml
Type: WordDocument
Parameter Sets: Paragraph
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
{{ Fill Italic Description }}

```yaml
Type: Nullable`1[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Paragraph
{{ Fill Paragraph Description }}

```yaml
Type: WordParagraph
Parameter Sets: Paragraph
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReturnObject
{{ Fill ReturnObject Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Style
{{ Fill Style Description }}

```yaml
Type: WordParagraphStyles
Parameter Sets: (All)
Aliases:
Accepted values: Normal, Heading1, Heading2, Heading3, Heading4, Heading5, Heading6, Heading7, Heading8, Heading9, ListParagraph, Custom

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
{{ Fill Text Description }}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
{{ Fill Underline Description }}

```yaml
Type: Nullable`1[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ProgressAction
{{ Fill ProgressAction Description }}

```yaml
Type: ActionPreference
Parameter Sets: (All)
Aliases: proga

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### None

## OUTPUTS

### System.Object
## NOTES

## RELATED LINKS
