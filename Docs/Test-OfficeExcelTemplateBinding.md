---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Test-OfficeExcelTemplateBinding
## SYNOPSIS
Validates Excel template markers against supplied bindings before applying a template.

## SYNTAX
### Path (Default)
```powershell
Test-OfficeExcelTemplateBinding [-InputPath] <string> -Binding <IDictionary> [-Quiet] [-AsMarkdown] [-ThrowOnMissing] [<CommonParameters>]
```

### Document
```powershell
Test-OfficeExcelTemplateBinding -Document <ExcelDocument> -Binding <IDictionary> [-Quiet] [-AsMarkdown] [-ThrowOnMissing] [<CommonParameters>]
```

## DESCRIPTION
Validates Excel template markers against supplied bindings before applying a template.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $bindings = @{
    CustomerName = 'Northwind'
    InvoiceDate  = Get-Date
    Total        = 1250.75
}
$result = Test-OfficeExcelTemplateBinding -Path .\InvoiceTemplate.xlsx -Binding $bindings -AsMarkdown
$result | Set-Content .\InvoiceTemplateBinding.md
```

Uses OfficeIMO template inspection and returns either structured missing-marker data or the reusable Markdown marker report.

## PARAMETERS

### -AsMarkdown
Return OfficeIMO's Markdown marker report.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Binding
Template bindings keyed by marker name.

```yaml
Type: IDictionary
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Workbook document.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -InputPath
Workbook path.

```yaml
Type: String
Parameter Sets: Path
Aliases: Path, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Quiet
Return only a Boolean pass/fail value.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ThrowOnMissing
Throw when any marker is missing.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
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

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Management.Automation.PSObject
System.String
System.Boolean`

## RELATED LINKS

- None
