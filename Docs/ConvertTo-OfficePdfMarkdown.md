---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficePdfMarkdown
## SYNOPSIS
Converts PDF logical text readback to Markdown.

## SYNTAX
### __AllParameterSets
```powershell
ConvertTo-OfficePdfMarkdown [-Path] <string> [-PageRange <string>] [-OutputPath <string>] [<CommonParameters>]
```

## DESCRIPTION
Converts PDF logical text readback to Markdown.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ConvertTo-OfficePdfMarkdown -Path .\Examples\Documents\Report.pdf -PageRange '1-3' -OutputPath .\Examples\Documents\Report.md
            Get-Content .\Examples\Documents\Report.md -TotalCount 20
```

Writes Markdown readback for selected pages to a file.

## PARAMETERS

### -OutputPath
Optional output Markdown file path.

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

### -PageRange
Optional page ranges such as 1-3,5.

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

### -Path
PDF file path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `System.String
System.IO.FileInfo`

## RELATED LINKS

- None
