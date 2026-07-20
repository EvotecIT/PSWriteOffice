---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeDocumentReader
## SYNOPSIS
Creates an immutable fully configured OfficeIMO document reader.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeDocumentReader [-ReaderAllOptions <ReaderAllOptions>] [-Processor <IOfficeDocumentProcessor[]>] [-OcrEngine <IOfficeOcrEngine>] [-TesseractOptions <TesseractOcrEngineOptions>] [-ProcessOcrOptions <ProcessOfficeOcrEngineOptions>] [-OcrOptions <OfficeDocumentOcrExecutionOptions>] [-UseTesseract] [-TesseractExecutablePath <string>] [-TesseractLanguage <string>] [-TesseractDataPath <string>] [-TesseractDpi <int>] [-TesseractTimeoutSeconds <int>] [-MaxStoreItems <int>] [-AllStoreItems] [-MaxConcurrentReads <int>] [-ProcessorFailureBehavior <OfficeDocumentProcessorFailureBehavior>] [<CommonParameters>]
```

## DESCRIPTION
Creates an immutable fully configured OfficeIMO document reader.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $reader = New-OfficeDocumentReader -TesseractLanguage 'eng+pol' -MaxStoreItems 5000 -ProcessorFailureBehavior ContinueWithDiagnostic
```

The returned reader can be supplied to every PSWriteOffice Reader command.

## PARAMETERS

### -AllStoreItems
Project every matching item from each email store.

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

### -MaxConcurrentReads
Maximum asynchronous reads allowed in flight.

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

### -MaxStoreItems
Maximum PST, OST, OLM, or EMLX items projected from each store. The default is 1,000.

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

### -OcrEngine
Caller-provided OCR engine.

```yaml
Type: IOfficeOcrEngine
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OcrOptions
Optional OCR execution limits and merge behavior.

```yaml
Type: OfficeDocumentOcrExecutionOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ProcessOcrOptions
Configure the generic JSON file-protocol OCR process adapter.

```yaml
Type: ProcessOfficeOcrEngineOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Processor
Additional ordered processors to run after document extraction.

```yaml
Type: IOfficeDocumentProcessor[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ProcessorFailureBehavior
Behavior when a processor fails.

```yaml
Type: OfficeDocumentProcessorFailureBehavior
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Throw, ContinueWithDiagnostic, StopWithDiagnostic

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ReaderAllOptions
Advanced format-specific settings supplied by a .NET host.

```yaml
Type: ReaderAllOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -TesseractDataPath
Optional Tesseract tessdata directory.

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

### -TesseractDpi
Optional input DPI passed to Tesseract.

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

### -TesseractExecutablePath
Tesseract executable path or command name.

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

### -TesseractLanguage
Tesseract language expression such as eng or eng+pol.

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

### -TesseractOptions
Configure the built-in Tesseract command-line OCR adapter.

```yaml
Type: TesseractOcrEngineOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -TesseractTimeoutSeconds
Maximum Tesseract process duration in seconds. The default is 120.

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

### -UseTesseract
Enable the built-in Tesseract command-line OCR adapter with default settings.

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

- `OfficeIMO.Reader.OfficeDocumentReader`

## RELATED LINKS

- None
