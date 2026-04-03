---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordTableCell
## SYNOPSIS
Enters a specific table cell and executes nested DSL content inside it.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeWordTableCell [-Row] <int> [-Column] <int> [[-Content] <scriptblock>] [-PassThru] [<CommonParameters>]
```

### Table
```powershell
Add-OfficeWordTableCell [-Table] <WordTable> [-Row] <int> [-Column] <int> [[-Content] <scriptblock>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Enters a specific table cell and executes nested DSL content inside it.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>WordTable -Data $Rows {
    WordTableCell -Row 1 -Column 0 {
        WordParagraph { WordText 'Details' }
        WordList {
            WordListItem 'One'
            WordListItem 'Two'
        }
    }
}
```

Targets the data cell at row 1, column 0 and writes nested content inside it.
You can mix paragraphs, images, lists, and nested tables inside the same cell.

## PARAMETERS

### -Column
Zero-based column index.

```yaml
Type: Int32
Parameter Sets: Context, Table
Aliases: None
Possible values: 

Required: True
Position: named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
DSL content executed inside the selected cell.

```yaml
Type: ScriptBlock
Parameter Sets: Context, Table
Aliases: None
Possible values: 

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the selected WordTableCell.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Table
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Row
Zero-based row index.

```yaml
Type: Int32
Parameter Sets: Context, Table
Aliases: None
Possible values: 

Required: True
Position: named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -Table
Optional table to target outside the active DSL table scope.

```yaml
Type: WordTable
Parameter Sets: Table
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordTable`

## OUTPUTS

- `OfficeIMO.Word.WordTableCell`

## RELATED LINKS

- None
