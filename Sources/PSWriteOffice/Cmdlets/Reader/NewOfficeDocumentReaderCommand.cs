using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using OfficeIMO.Reader.Ocr.Process;
using OfficeIMO.Reader.Ocr.Tesseract;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Creates an immutable fully configured OfficeIMO document reader.</summary>
/// <example>
///   <summary>Create a reader with OCR and a resilient processor policy.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ocr = [OfficeIMO.Reader.Ocr.Tesseract.TesseractOcrEngineOptions]::new(); $ocr.Language = 'eng+pol'; $reader = New-OfficeDocumentReader -TesseractOptions $ocr -ProcessorFailureBehavior ContinueWithDiagnostic</code>
///   <para>The returned reader can be supplied to every PSWriteOffice Reader command.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeDocumentReader")]
[OutputType(typeof(OfficeDocumentReader))]
public sealed class NewOfficeDocumentReaderCommand : PSCmdlet
{
    /// <summary>Optional format-specific settings captured while OfficeIMO Reader handlers are registered.</summary>
    [Parameter]
    public ReaderAllOptions? ReaderAllOptions { get; set; }

    /// <summary>Additional ordered processors to run after document extraction.</summary>
    [Parameter]
    public IOfficeDocumentProcessor[]? Processor { get; set; }

    /// <summary>Caller-provided OCR engine.</summary>
    [Parameter]
    public IOfficeOcrEngine? OcrEngine { get; set; }

    /// <summary>Configure the built-in Tesseract command-line OCR adapter.</summary>
    [Parameter]
    public TesseractOcrEngineOptions? TesseractOptions { get; set; }

    /// <summary>Configure the generic JSON file-protocol OCR process adapter.</summary>
    [Parameter]
    public ProcessOfficeOcrEngineOptions? ProcessOcrOptions { get; set; }

    /// <summary>Optional OCR execution limits and merge behavior.</summary>
    [Parameter]
    public OfficeDocumentOcrExecutionOptions? OcrOptions { get; set; }

    /// <summary>Maximum asynchronous reads allowed in flight.</summary>
    [Parameter]
    [ValidateRange(1, 64)]
    public int? MaxConcurrentReads { get; set; }

    /// <summary>Behavior when a processor fails.</summary>
    [Parameter]
    public OfficeDocumentProcessorFailureBehavior ProcessorFailureBehavior { get; set; } = OfficeDocumentProcessorFailureBehavior.Throw;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var configuredOcr = new List<IOfficeOcrEngine>();
        if (OcrEngine != null) configuredOcr.Add(OcrEngine);
        if (TesseractOptions != null) configuredOcr.Add(new TesseractOcrEngine(TesseractOptions));
        if (ProcessOcrOptions != null) configuredOcr.Add(new ProcessOfficeOcrEngine(ProcessOcrOptions));
        if (configuredOcr.Count > 1)
        {
            throw new PSArgumentException("Specify only one of -OcrEngine, -TesseractOptions, or -ProcessOcrOptions.");
        }

        var builder = ReaderCommandUtilities.CreateBuilder(ReaderAllOptions)
            .WithProcessorFailureBehavior(ProcessorFailureBehavior);
        if (MaxConcurrentReads.HasValue) builder.WithMaxConcurrentReads(MaxConcurrentReads.Value);
        if (Processor != null) builder.AddProcessors(Processor.Where(static value => value != null));
        if (configuredOcr.Count == 1) builder.AddProcessor(new OfficeDocumentOcrProcessor(configuredOcr[0], OcrOptions));
        WriteObject(builder.Build());
    }
}
