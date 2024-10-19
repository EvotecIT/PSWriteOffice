---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version:
schema: 2.0.0
---

# Import-OfficeExcel

## SYNOPSIS
Provides a way to converting an Excel file into PowerShell objects.

## SYNTAX

```
Import-OfficeExcel [-FilePath] <String> [[-WorkSheetName] <String[]>] [-ProgressAction <ActionPreference>]
 [<CommonParameters>]
```

## DESCRIPTION
Provides a way to converting an Excel file into PowerShell objects.
If Worksheet is not specified, all worksheets will be imported and returned as a hashtable of worksheet names and worksheet objects.
If Worksheet is specified, only the specified worksheet will be imported and returned as an array of PSCustomObjects

## EXAMPLES

### EXAMPLE 1
```
$FilePath = "$PSScriptRoot\Documents\Test5.xlsx"
```

$ImportedData1 = Import-OfficeExcel -FilePath $FilePath
$ImportedData1 | Format-Table

### EXAMPLE 2
```
$FilePath = "$PSScriptRoot\Documents\Excel.xlsx"
```

$ImportedData2 = Import-OfficeExcel -FilePath $FilePath -WorkSheetName 'Contact3'
$ImportedData2 | Format-Table

## PARAMETERS

### -FilePath
The path to the Excel file to import.

```yaml
Type: String
Parameter Sets: (All)
Aliases: LiteralPath

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WorkSheetName
The name of the worksheet to import.
If not specified, all worksheets will be imported.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ProgressAction
{{ Fill ProgressAction Description }}

```yaml
Type: ActionPreference
Parameter Sets: (All)
Aliases: proga

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES
General notes

## RELATED LINKS
