---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWord
## SYNOPSIS
Opens an existing Word document.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeWord [-InputPath] <string> [-ReadOnly] [-AutoSave] [<CommonParameters>]
```

## DESCRIPTION
Opens an existing Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>$doc = Get-OfficeWord -Path .\Report.docx -ReadOnly
```

Loads Report.docx and exposes the document object for querying.

## PARAMETERS

### -AutoSave
Enable AutoSave when editing.

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

### -InputPath
Path to the .docx. Accepts PS paths.

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

### -ReadOnly
Open in read-only mode.

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

- `System.Object`

## RELATED LINKS

- None

