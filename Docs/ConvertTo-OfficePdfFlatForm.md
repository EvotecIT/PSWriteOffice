---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficePdfFlatForm
## SYNOPSIS
Converts a PDF with simple AcroForm fields into a flat PDF.

## SYNTAX
### __AllParameterSets
```powershell
ConvertTo-OfficePdfFlatForm [-Path] <string> [-OutputPath] <string> [<CommonParameters>]
```

## DESCRIPTION
Converts a PDF with simple AcroForm fields into a flat PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
ConvertTo-OfficePdfFlatForm -Path 'C:\Path'
```


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
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
