---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordSection
## SYNOPSIS
Adds or reuses a section inside the current Word document.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordSection [[-Content] <scriptblock>] [-BreakType <SectionMarkValues>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds or reuses a section inside the current Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>New-OfficeWord -Path .\doc.docx { Add-OfficeWordSection { Add-OfficeWordParagraph -Text 'Hello' } }
```

Creates a document and inserts a section that contains a single paragraph.

## PARAMETERS

### -BreakType
Optional section break type.

```yaml
Type: Nullable`1
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Content
DSL scriptblock executed within the section scope.

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

### -PassThru
Emit the created WordSection.

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

