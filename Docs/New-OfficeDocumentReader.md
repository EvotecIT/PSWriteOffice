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
New-OfficeDocumentReader [-ReaderAllOptions <ReaderAllOptions>] [-Processor <IOfficeDocumentProcessor[]>] [-OcrEngine <IOfficeOcrEngine>] [-TesseractOptions <TesseractOcrEngineOptions>] [-ProcessOcrOptions <ProcessOfficeOcrEngineOptions>] [-OcrOptions <OfficeDocumentOcrExecutionOptions>] [-MaxConcurrentReads <int>] [-ProcessorFailureBehavior <OfficeDocumentProcessorFailureBehavior>] [<CommonParameters>]
```

## DESCRIPTION
Creates an immutable fully configured OfficeIMO document reader.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $ocr = [OfficeIMO.Reader.Ocr.Tesseract.TesseractOcrEngineOptions]::new(); $ocr.Language = 'eng+pol'; $reader = New-OfficeDocumentReader -TesseractOptions $ocr -ProcessorFailureBehavior ContinueWithDiagnostic
```

The returned reader can be supplied to every PSWriteOffice Reader command.

## PARAMETERS

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
Optional format-specific settings captured while OfficeIMO Reader handlers are registered.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Reader.OfficeDocumentReader`

## RELATED LINKS

- None
