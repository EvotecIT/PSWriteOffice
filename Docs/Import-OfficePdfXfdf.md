---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Import-OfficePdfXfdf
## SYNOPSIS
Imports bounded DTD-free XFDF through the validated PDF form filler.

## SYNTAX
### Text (Default)
```powershell
Import-OfficePdfXfdf [-Path] <string> -Xfdf <string> -OutputPath <string> [-Options <PdfFormFillerOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### File
```powershell
Import-OfficePdfXfdf [-Path] <string> -XfdfPath <string> -OutputPath <string> [-Options <PdfFormFillerOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Imports bounded DTD-free XFDF through the validated PDF form filler.

## EXAMPLES

### EXAMPLE 1
```powershell
Import-OfficePdfXfdf -Xfdf 'Value' -OutputPath 'C:\Path'
```


### EXAMPLE 2
```powershell
Import-OfficePdfXfdf -XfdfPath 'C:\Path' -OutputPath 'C:\Path'
```


## PARAMETERS

### -Options
Optional validated form filling behavior.

```yaml
Type: PdfFormFillerOptions
Parameter Sets: Text, File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OutputPath
Destination PDF path.

```yaml
Type: String
Parameter Sets: Text, File
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Return the rewritten fluent PDF document.

```yaml
Type: SwitchParameter
Parameter Sets: Text, File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
Source PDF path.

```yaml
Type: String
Parameter Sets: Text, File
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Xfdf
XFDF XML.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -XfdfPath
Path to an XFDF file.

```yaml
Type: String
Parameter Sets: File
Aliases: None
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

- `System.String`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
