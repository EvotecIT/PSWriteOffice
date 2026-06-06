---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Move-OfficePdfPage
## SYNOPSIS
Moves selected pages before another page and writes a new PDF.

## SYNTAX
### __AllParameterSets
```powershell
Move-OfficePdfPage -Path <string> -PageRange <string> -BeforePage <int> -OutputPath <string> [<CommonParameters>]
```

## DESCRIPTION
Moves selected pages before another page and writes a new PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
Move-OfficePdfPage -Path 'C:\Path' -PageRange 'Value' -BeforePage 1 -OutputPath 'C:\Path'
```


## PARAMETERS

### -BeforePage
One-based page number before which selected pages are inserted. Use page count + 1 to move to the end.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OutputPath
Output PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PageRange
Page ranges such as 1-3,5.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
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
