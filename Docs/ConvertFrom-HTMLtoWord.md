---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version:
schema: 2.0.0
---

# ConvertFrom-HTMLtoWord

## SYNOPSIS
Converts HTML input to Microsoft Word Document

## SYNTAX

### HTML (Default)
```
ConvertFrom-HTMLtoWord -OutputFile <String> -SourceHTML <String> [-Show] [<CommonParameters>]
```

### HTMLFile
```
ConvertFrom-HTMLtoWord -OutputFile <String> -FileHTML <String> [-Show] [<CommonParameters>]
```

## DESCRIPTION
Converts HTML input to Microsoft Word Document

## EXAMPLES

### EXAMPLE 1
```
$Objects = @(
```

\[PSCustomObject\] @{ Test = 1; Test2 = 'Test'; Test3 = 'Ok' }
    \[PSCustomObject\] @{ Test = 1; Test2 = 'Test'; Test3 = 'Ok' }
)

New-HTML {
    New-HTMLText -Text 'This is a test', ' another test' -FontSize 30pt
    New-HTMLTable -DataTable $Objects -Simplify
} -Online -FilePath $PSScriptRoot\Documents\Test.html

ConvertFrom-HTMLToWord -OutputFile $PSScriptRoot\Documents\TestHTML.docx -FileHTML $PSScriptRoot\Documents\Test.html -Show

### EXAMPLE 2
```
$Objects = @(
```

\[PSCustomObject\] @{ Test = 1; Test2 = 'Test'; Test3 = 'Ok' }
\[PSCustomObject\] @{ Test = 1; Test2 = 'Test'; Test3 = 'Ok' }
)

$Test = New-HTML {
    New-HTMLText -Text 'This is a test', ' another test' -FontSize 30pt
    New-HTMLTable -DataTable $Objects -simplify
} -Online

ConvertFrom-HTMLToWord -OutputFile $PSScriptRoot\Documents\TestHTML.docx -HTML $Test -Show

## PARAMETERS

### -OutputFile
Path to the file to save converted HTML

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FileHTML
Input HTML loaded straight from file

```yaml
Type: String
Parameter Sets: HTMLFile
Aliases: InputFile

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceHTML
Input HTML loaded from string

```yaml
Type: String
Parameter Sets: HTML
Aliases: HTML

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Once conversion ends show the resulting document

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: True
Position: Named
Default value: False
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
