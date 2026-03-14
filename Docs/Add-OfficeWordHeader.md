---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordHeader
## SYNOPSIS
Adds content to a section header.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordHeader [[-Content] <scriptblock>] [-Type <HeaderFooterValues>] [<CommonParameters>]
```

## DESCRIPTION
Adds content to a section header.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Add-OfficeWordSection { Add-OfficeWordHeader { Add-OfficeWordParagraph -Text 'Confidential' -Style Heading3 } }
```

Creates a section header that prints “Confidential”.

## PARAMETERS

### -Content
DSL scriptblock to execute inside the header.

```yaml
Type: ScriptBlock
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Type
The header type to modify.

```yaml
Type: HeaderFooterValues
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

