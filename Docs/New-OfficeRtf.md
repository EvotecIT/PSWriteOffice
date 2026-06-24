---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeRtf
## SYNOPSIS
Creates an RTF document with plain paragraph content.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeRtf [-OutputPath] <string> [[-Text] <string[]>] [-PassThru] [-NoSave] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Creates an RTF document with plain paragraph content.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $file = New-OfficeRtf -Path .\Report.rtf -Text 'Summary', 'Ready for review' -PassThru
Get-OfficeRtf -Path $file.FullName
```

Creates an RTF document with two paragraphs and returns the file.

## PARAMETERS

### -NoSave
Return the OfficeIMO RTF document without saving.

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
Destination path for the RTF file.

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

### -Text
Plain paragraph text to add to the document.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`

## OUTPUTS

- `System.IO.FileInfo
OfficeIMO.Rtf.RtfDocument`

## RELATED LINKS

- None
