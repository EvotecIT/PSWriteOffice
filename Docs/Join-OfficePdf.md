---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Join-OfficePdf
## SYNOPSIS
Joins multiple PDF files into a single PDF.

## SYNTAX
### __AllParameterSets
```powershell
Join-OfficePdf [-Path] <string[]> [-OutputPath] <string> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Joins multiple PDF files into a single PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $cover = '.\Examples\Documents\Cover.pdf'
            $report = '.\Examples\Documents\Report.pdf'
            Join-OfficePdf -Path $cover, $report -OutputPath .\Examples\Documents\Combined.pdf -PassThru
            Get-OfficePdfInfo -Path .\Examples\Documents\Combined.pdf | Select-Object PageCount
```

Writes a single PDF containing the input documents in the requested order, then checks the result.

## PARAMETERS

### -OutputPath
Output PDF path.

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

### -PassThru
Emit the saved file.

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

### -Path
Input PDF paths in output order.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
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

- `System.IO.FileInfo`

## RELATED LINKS

- None
