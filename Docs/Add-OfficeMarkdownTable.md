---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeMarkdownTable
## SYNOPSIS
Adds a Markdown table from objects.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeMarkdownTable [-InputObject <Object>] [-DisableAutoAlign] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeMarkdownTable -Document <MarkdownDoc> [-InputObject <Object>] [-DisableAutoAlign] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a Markdown table from objects.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>MarkdownTable -InputObject $rows
```

Appends a Markdown table using the supplied objects.

### EXAMPLE 2
```powershell
PS>$doc = New-OfficeMarkdown -Path .\Report.md -NoSave -PassThru
$doc | MarkdownTable -InputObject $summary -PassThru | MarkdownTable -InputObject $details
```

Creates two tables in sequence within the same Markdown document.

## PARAMETERS

### -DisableAutoAlign
Disable automatic alignment heuristics for tables.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Markdown document to update outside the DSL context.

```yaml
Type: MarkdownDoc
Parameter Sets: Document
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -InputObject
Objects to convert into a Markdown table.

```yaml
Type: Object
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -PassThru
Emit the Markdown document after appending the table.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
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

- `OfficeIMO.Markdown.MarkdownDoc
System.Object`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## RELATED LINKS

- None

