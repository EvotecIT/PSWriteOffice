---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordParagraph
## SYNOPSIS
Adds a paragraph to the current section/header/footer context.

## SYNTAX
### Text (Default)
```powershell
Add-OfficeWordParagraph [[-Text] <string>] [-Alignment <JustificationValues>] [-Style <WordParagraphStyles>] [-PassThru] [<CommonParameters>]
```

### Content
```powershell
Add-OfficeWordParagraph [[-Content] <scriptblock>] [-Text <string>] [-Alignment <JustificationValues>] [-Style <WordParagraphStyles>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a paragraph to the current section/header/footer context.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Add-OfficeWordParagraph { Add-OfficeWordText -Text 'Hello '; Add-OfficeWordText -Text 'World' -Bold }
```

Outputs “Hello World” with the second word bolded.

## PARAMETERS

### -Alignment
Paragraph justification.

```yaml
Type: Nullable`1
Parameter Sets: Text, Content
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Content
Nested DSL content (runs, lists, images).

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values: 

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the WordParagraph for further use.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Content
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Style
Paragraph style.

```yaml
Type: Nullable`1
Parameter Sets: Text, Content
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Text
Optional initial paragraph text.

```yaml
Type: String
Parameter Sets: Text, Content
Aliases: None
Possible values: 

Required: False
Position: 0
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

