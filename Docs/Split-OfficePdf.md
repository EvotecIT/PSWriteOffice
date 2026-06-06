---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Split-OfficePdf
## SYNOPSIS
Splits a PDF into one file per page.

## SYNTAX
### __AllParameterSets
```powershell
Split-OfficePdf [-Path] <string> [-OutputDirectory] <string> [-Prefix <string>] [<CommonParameters>]
```

## DESCRIPTION
Splits a PDF into one file per page.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $pages = Split-OfficePdf -Path .\Examples\Documents\Combined.pdf -OutputDirectory .\Examples\Documents\Pages -Prefix 'combined-page'
            $pages | Select-Object Name, Length
```

Creates one output PDF for each page and returns the written files.

## PARAMETERS

### -OutputDirectory
Output directory.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
Input PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Prefix
Output file prefix.

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

- `System.IO.FileInfo`

## RELATED LINKS

- None
