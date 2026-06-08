---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePdfPage
## SYNOPSIS
Sets page-level PDF properties and writes a new PDF.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficePdfPage -Path <string> -Rotation <int> -OutputPath <string> [-PageRange <string>] [<CommonParameters>]
```

## DESCRIPTION
Sets page-level PDF properties and writes a new PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $proof = @(
    Set-OfficePdfPage -Path .\Examples\Documents\Scanned.pdf -PageRange '2,4' -Rotation 90 -OutputPath .\Examples\Documents\Scanned-Rotated.pdf
    Get-OfficePdfInfo -Path .\Examples\Documents\Scanned-Rotated.pdf | Select-Object PageCount
)
$proof
```

Rotates selected pages and writes a new PDF.

## PARAMETERS

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
Page ranges such as 1-3,5. Omit to affect all pages.

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

### -Rotation
Rotation in degrees. Supported values are 0, 90, 180, and 270.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 0, 90, 180, 270

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
