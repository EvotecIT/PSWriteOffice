---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeMarkdownTaskList
## SYNOPSIS
Adds a Markdown task list.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeMarkdownTaskList [-Items] <string[]> [-Completed <int[]>] [-AllCompleted] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeMarkdownTaskList [-Items] <string[]> -Document <MarkdownDoc> [-Completed <int[]>] [-AllCompleted] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a Markdown task list.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>MarkdownTaskList -Items 'Draft','Review','Ship' -Completed 1
```

Appends an unordered task list and marks the selected items as completed.

## PARAMETERS

### -AllCompleted
Mark every task as completed.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Completed
Zero-based item indexes that should be marked complete.

```yaml
Type: Int32[]
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Markdown document to update outside the DSL context.

```yaml
Type: MarkdownDoc
Parameter Sets: Document
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Items
Task text entries to include in the checklist.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the updated Markdown document.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
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

- `OfficeIMO.Markdown.MarkdownDoc`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## RELATED LINKS

- None

