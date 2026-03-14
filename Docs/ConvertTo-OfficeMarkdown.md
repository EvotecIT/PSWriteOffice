---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficeMarkdown
## SYNOPSIS
Converts objects into a Markdown table.

## SYNTAX
### __AllParameterSets
```powershell
ConvertTo-OfficeMarkdown [-InputObject <Object>] [-DisableAutoAlign] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Converts objects into a Markdown table.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>$markdown = $data | ConvertTo-OfficeMarkdown
```

Generates Markdown table text from the input objects.

### EXAMPLE 2
```powershell
PS>$doc = $data | ConvertTo-OfficeMarkdown -PassThru
$doc.P('Totals above'); $doc.ToMarkdown()
```

Builds a table and appends more content using the MarkdownDoc API.

### EXAMPLE 3
```powershell
PS>$markdown = $data | ConvertTo-OfficeMarkdown -DisableAutoAlign
```

Forces left-aligned columns instead of auto-aligned output.

## PARAMETERS

### -DisableAutoAlign
Disable automatic alignment heuristics for tables.

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

### -InputObject
Objects to convert into Markdown.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -PassThru
Emit a Markdown document object instead of text.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.Object`

## OUTPUTS

- `System.String
OfficeIMO.Markdown.MarkdownDoc`

## RELATED LINKS

- None

