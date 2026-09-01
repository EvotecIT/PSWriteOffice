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

    /// <summary>Friendly OCR languages. Supply more than one value to recognize multilingual content.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public OfficeOcrLanguage[]? Language { get; set; }

    /// <summary>Advanced raw Tesseract expression for caller-installed custom trained-data models.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public string? TesseractLanguageExpression { get; set; }

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
            Languages = source.Languages,
            CustomLanguageExpression = source.CustomLanguageExpression,
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
            effective.Languages = CombineLanguages(Language);
            effective.CustomLanguageExpression = null;
            effective.Tesseract.Language = "eng";
        }

        if (MyInvocation.BoundParameters.ContainsKey(nameof(TesseractLanguageExpression)))
        {
            effective.Languages = OfficeOcrLanguage.English;
            effective.CustomLanguageExpression = TesseractLanguageExpression;
            effective.Tesseract.Language = "eng";
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

    private OfficeOcrLanguage CombineLanguages(OfficeOcrLanguage[]? languages)
    {
        if (languages == null || languages.Length == 0)
        {
            throw new PSArgumentException("Select at least one OCR language.", nameof(Language));
        }
        if (MyInvocation.BoundParameters.ContainsKey(nameof(TesseractLanguageExpression)))
        {
            throw new PSArgumentException(
                "Use -Language or the advanced -TesseractLanguageExpression parameter, not both.",
                nameof(Language));
        }

        OfficeOcrLanguage combined = 0;
        foreach (OfficeOcrLanguage language in languages)
        {
            combined |= language;
        }
        _ = combined.ToTesseractExpression();
        return combined;
    }
}
