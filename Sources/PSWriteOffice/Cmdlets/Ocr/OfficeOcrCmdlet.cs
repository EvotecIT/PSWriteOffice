using System.Management.Automation;
using OfficeIMO.Reader.Ocr;
using OfficeIMO.Reader.Ocr.Tesseract;

namespace PSWriteOffice.Cmdlets.Ocr;

/// <summary>Common easy-OCR options shared by PSWriteOffice OCR commands.</summary>
public abstract class OfficeOcrCmdlet : AsyncPSCmdlet
{
    /// <summary>Advanced OfficeIMO OCR options. Convenience parameters override matching values.</summary>
    [Parameter]
    public OfficeOcrOptions? Options { get; set; }

    /// <summary>Tesseract language expression, such as eng or eng+pol. The default is eng.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public string? Language { get; set; }

    /// <summary>Explicit Tesseract executable path. By default OfficeIMO securely discovers an installed runtime.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public string? TesseractPath { get; set; }

    /// <summary>Explicit directory containing Tesseract trained-data files.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public string? TessdataDirectory { get; set; }

    /// <summary>Do not download checksum-pinned curated language data when a requested language is missing.</summary>
    [Parameter]
    public SwitchParameter NoLanguageDownload { get; set; }

    /// <summary>Builds an independent options snapshot and applies bound convenience parameters.</summary>
    protected OfficeOcrOptions CreateOptions()
    {
        OfficeOcrOptions source = Options ?? new OfficeOcrOptions();
        if (source.Tesseract == null)
        {
            throw new PSArgumentException("Options.Tesseract cannot be null.", nameof(Options));
        }

        if (source.Pdf == null)
        {
            throw new PSArgumentException("Options.Pdf cannot be null.", nameof(Options));
        }

        if (source.LanguageData == null)
        {
            throw new PSArgumentException("Options.LanguageData cannot be null.", nameof(Options));
        }

        var effective = new OfficeOcrOptions
        {
            OutputConflictPolicy = source.OutputConflictPolicy,
            Tesseract = source.Tesseract.Clone(),
            Pdf = source.Pdf.Clone(),
            LanguageData = new TesseractLanguageDataOptions
            {
                CacheDirectory = source.LanguageData.CacheDirectory,
                HttpClient = source.LanguageData.HttpClient,
                MaxBytesPerLanguage = source.LanguageData.MaxBytesPerLanguage
            },
            ProvisionMissingLanguageData = source.ProvisionMissingLanguageData
        };

        if (MyInvocation.BoundParameters.ContainsKey(nameof(Language)))
        {
            effective.Tesseract.Language = Language;
        }

        if (MyInvocation.BoundParameters.ContainsKey(nameof(TesseractPath)))
        {
            effective.Tesseract.ExecutablePath = TesseractPath!;
        }

        if (MyInvocation.BoundParameters.ContainsKey(nameof(TessdataDirectory)))
        {
            effective.Tesseract.TessdataDirectory = TessdataDirectory;
        }

        if (NoLanguageDownload.IsPresent)
        {
            effective.ProvisionMissingLanguageData = false;
        }

        return effective;
    }
}
