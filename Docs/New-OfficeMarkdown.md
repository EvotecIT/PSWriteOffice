---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeMarkdown
## SYNOPSIS
Creates a Markdown document using a DSL scriptblock.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeMarkdown [-OutputPath] <string> [[-Content] <scriptblock>] [-PassThru] [-NoSave] [<CommonParameters>]
```

## DESCRIPTION
Creates a Markdown document using a DSL scriptblock.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>New-OfficeMarkdown -Path .\README.md { MarkdownHeading -Level 1 -Text 'Report'; MarkdownTable -InputObject $data }
```

Creates a README file with a heading and table content.

### EXAMPLE 2
```powershell
PS>New-OfficeMarkdown -Path .\Report.md {
MarkdownHeading -Level 1 -Text 'Summary'
MarkdownTable -InputObject $summary
MarkdownHeading -Level 2 -Text 'Details'
MarkdownTable -InputObject $details
}
```

Creates a report with two tables separated by headings.

## PARAMETERS

### -Content
DSL scriptblock describing Markdown content.

```yaml
Type: ScriptBlock
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoSave
Skip saving after executing the DSL.

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

### -OutputPath
Destination path for the Markdown file.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath, Path
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit a FileInfo for chaining.

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

- `None`

## OUTPUTS

- `System.IO.FileInfo
OfficeIMO.Markdown.MarkdownDoc`

## RELATED LINKS

- None

