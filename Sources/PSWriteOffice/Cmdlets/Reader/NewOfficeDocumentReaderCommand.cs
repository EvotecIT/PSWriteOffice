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
///   <code>$reader = New-OfficeDocumentReader -TesseractLanguage 'eng+pol' -MaxStoreItems 5000 -ProcessorFailureBehavior ContinueWithDiagnostic</code>
///   <para>The returned reader can be supplied to every PSWriteOffice Reader command.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeDocumentReader")]
[OutputType(typeof(OfficeDocumentReader))]
public sealed class NewOfficeDocumentReaderCommand : PSCmdlet
{
    /// <summary>Advanced format-specific settings supplied by a .NET host.</summary>
    [Parameter(DontShow = true)]
    public ReaderAllOptions? ReaderAllOptions { get; set; }

    /// <summary>Additional ordered processors to run after document extraction.</summary>
    [Parameter(DontShow = true)]
    public IOfficeDocumentProcessor[]? Processor { get; set; }

    /// <summary>Caller-provided OCR engine.</summary>
    [Parameter(DontShow = true)]
    public IOfficeOcrEngine? OcrEngine { get; set; }

    /// <summary>Configure the built-in Tesseract command-line OCR adapter.</summary>
    [Parameter(DontShow = true)]
    public TesseractOcrEngineOptions? TesseractOptions { get; set; }

    /// <summary>Configure the generic JSON file-protocol OCR process adapter.</summary>
    [Parameter(DontShow = true)]
    public ProcessOfficeOcrEngineOptions? ProcessOcrOptions { get; set; }

    /// <summary>Optional OCR execution limits and merge behavior.</summary>
    [Parameter(DontShow = true)]
    public OfficeDocumentOcrExecutionOptions? OcrOptions { get; set; }

    /// <summary>Enable the built-in Tesseract command-line OCR adapter with default settings.</summary>
    [Parameter]
    public SwitchParameter UseTesseract { get; set; }

    /// <summary>Tesseract executable path or command name.</summary>
    [Parameter]
    public string? TesseractExecutablePath { get; set; }

    /// <summary>Tesseract language expression such as eng or eng+pol.</summary>
    [Parameter]
    public string? TesseractLanguage { get; set; }

    /// <summary>Optional Tesseract tessdata directory.</summary>
    [Parameter]
    public string? TesseractDataPath { get; set; }

    /// <summary>Optional input DPI passed to Tesseract.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? TesseractDpi { get; set; }

    /// <summary>Maximum Tesseract process duration in seconds. The default is 120.</summary>
    [Parameter]
    [ValidateRange(1, 86400)]
    public int? TesseractTimeoutSeconds { get; set; }

    /// <summary>Maximum PST, OST, OLM, or EMLX items projected from each store. The default is 1,000.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxStoreItems { get; set; }

    /// <summary>Project every matching item from each email store.</summary>
    [Parameter]
    public SwitchParameter AllStoreItems { get; set; }

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
        ValidatePowerShellParameters();
        var configuredOcr = new List<IOfficeOcrEngine>();
        if (OcrEngine != null) configuredOcr.Add(OcrEngine);
        if (TesseractOptions != null) configuredOcr.Add(new TesseractOcrEngine(TesseractOptions));
        if (HasPowerShellTesseractConfiguration())
        {
            var tesseract = new TesseractOcrEngineOptions();
            if (!string.IsNullOrWhiteSpace(TesseractExecutablePath)) tesseract.ExecutablePath = TesseractExecutablePath!;
            if (!string.IsNullOrWhiteSpace(TesseractLanguage)) tesseract.Language = TesseractLanguage!;
            if (!string.IsNullOrWhiteSpace(TesseractDataPath)) tesseract.TessdataDirectory = TesseractDataPath!;
            if (TesseractDpi.HasValue) tesseract.Dpi = TesseractDpi.Value;
            if (TesseractTimeoutSeconds.HasValue) tesseract.Timeout = TimeSpan.FromSeconds(TesseractTimeoutSeconds.Value);
            configuredOcr.Add(new TesseractOcrEngine(tesseract));
        }
        if (ProcessOcrOptions != null) configuredOcr.Add(new ProcessOfficeOcrEngine(ProcessOcrOptions));
        if (configuredOcr.Count > 1)
        {
            throw new PSArgumentException("Specify only one of -OcrEngine, -TesseractOptions, or -ProcessOcrOptions.");
        }

        int? storeItemLimit = AllStoreItems.IsPresent ? int.MaxValue : MaxStoreItems;
        var powerShellConfiguration = ReaderCommandUtilities.BuildSearchConfiguration(
            includePageLocations: false,
            storeItemLimit);
        var builder = ReaderCommandUtilities.CreateBuilder(
                ReaderAllOptions ?? powerShellConfiguration.HandlerOptions)
            .WithProcessorFailureBehavior(ProcessorFailureBehavior);
        if (MaxConcurrentReads.HasValue) builder.WithMaxConcurrentReads(MaxConcurrentReads.Value);
        if (Processor != null) builder.AddProcessors(Processor.Where(static value => value != null));
        if (configuredOcr.Count == 1) builder.AddProcessor(new OfficeDocumentOcrProcessor(configuredOcr[0], OcrOptions));
        WriteObject(builder.Build());
    }

    private bool HasPowerShellTesseractConfiguration()
    {
        return UseTesseract.IsPresent ||
               !string.IsNullOrWhiteSpace(TesseractExecutablePath) ||
               !string.IsNullOrWhiteSpace(TesseractLanguage) ||
               !string.IsNullOrWhiteSpace(TesseractDataPath) ||
               TesseractDpi.HasValue ||
               TesseractTimeoutSeconds.HasValue;
    }

    private void ValidatePowerShellParameters()
    {
        if (AllStoreItems.IsPresent && MaxStoreItems.HasValue)
        {
            throw new PSArgumentException("Specify either -MaxStoreItems or -AllStoreItems, not both.");
        }
        if (ReaderAllOptions != null && (AllStoreItems.IsPresent || MaxStoreItems.HasValue))
        {
            throw new PSArgumentException(
                "Use -MaxStoreItems/-AllStoreItems or the advanced -ReaderAllOptions object, not both.");
        }
        if (TesseractOptions != null && HasPowerShellTesseractConfiguration())
        {
            throw new PSArgumentException(
                "Use PowerShell Tesseract parameters or the advanced -TesseractOptions object, not both.");
        }
    }
}
