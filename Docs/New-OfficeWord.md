---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeWord
## SYNOPSIS
Creates a Word document using the DSL.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeWord [-OutputPath] <string> [[-Content] <scriptblock>] [-TemplatePath <string>] [-PassThru] [-Open] [-NoSave] [-AutoSave] [-Password <string>] [-PdfPath <string>] [-PdfFontFamily <string>] [-PdfAllowSystemFontEmbedding] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Handles file creation or template cloning, scriptblock execution, optional autosave, and emits the document path when -PassThru is used.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\Report.docx { WordSection { WordParagraph 'Hello DSL' } } -Open
```

Builds a document, adds one paragraph, saves it to disk, and opens it.

### EXAMPLE 2
```powershell
PS> New-OfficeWord -TemplatePath .\Template.docx -Path .\Report.docx { WordParagraph -Text 'Generated content' -StyleId 'ReportBody' }
```

Copies the template to the output path, runs the DSL against the copied document, and saves it.

## PARAMETERS

### -AutoSave
Enable OfficeIMO AutoSave mode.

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

### -Content
DSL scriptblock describing document content.

```yaml
Type: ScriptBlock
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoSave
Skip saving after executing the DSL.

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

### -Open
Open the document after saving.

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
Destination path for the document.

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

### -Password
Password used to save the document as an encrypted package.

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

### -PdfAllowSystemFontEmbedding
Allow the native Word PDF converter to embed installed system fonts used by the document.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: AllowSystemFontEmbedding
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PdfFontFamily
Optional default font family used by the native Word PDF converter.

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

### -PdfPath
Optional PDF path to create from the same Word document before closing it.

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

### -TemplatePath
Existing .docx file to clone before running the DSL.

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

- `System.Object`

## RELATED LINKS

- None
