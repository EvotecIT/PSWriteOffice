using System.IO;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Ocr;

/// <summary>Creates a searchable PDF by adding invisible text from a discovered local OCR runtime.</summary>
/// <example>
///   <summary>Make a scanned PDF searchable.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficePdfSearchable -Path .\Scan.pdf -OutputPath .\Scan-Searchable.pdf</code>
///   <para>Preserves visible page content and writes geometry-aligned invisible English text.</para>
/// </example>
/// <example>
///   <summary>Recognize an English and Polish document and inspect the OCR evidence.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficePdfSearchable -Path .\Scan.pdf -OutputPath .\Searchable.pdf -Language English, Polish -PassThru</code>
///   <para>Returns recognition, filtering, page, provider, and model evidence instead of the output file.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficePdfSearchable", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
[OutputType(typeof(PdfSearchableOcrResult))]
public sealed class ConvertToOfficePdfSearchableCommand : OfficeOcrCmdlet
{
    /// <summary>Source PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Destination PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Overwrite an existing destination file.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Return the complete searchable-PDF OCR result instead of the output file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>PDF page render resolution used for recognition.</summary>
    [Parameter]
    [ValidateRange(72, 1200)]
    public double? RenderDpi { get; set; }

    /// <summary>Minimum normalized confidence accepted for searchable text.</summary>
    [Parameter]
    [ValidateRange(0.0, 1.0)]
    public double? MinimumConfidence { get; set; }

    /// <summary>Advanced PDF page selection, rendering, confidence, overlap, and resource limits.</summary>
    [Parameter]
    public PdfOcrMergeOptions? PdfOptions { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync()
    {
        string inputPath = PdfCommandUtilities.ResolveExistingFilePath(this, Path);
        string outputPath = PdfCommandUtilities.ResolveOutputFilePath(this, OutputPath, ".pdf", Force.IsPresent);
        if (!ShouldProcess(outputPath, "Create searchable PDF"))
        {
            return;
        }

        TesseractOcrSession session = await TesseractOcr
            .CreateSessionAsync(CreateSessionOptions(), CancelToken)
            .ConfigureAwait(false);
        PdfOcrMergeOptions options = PdfOptions?.Clone() ?? new PdfOcrMergeOptions();
        if (RenderDpi.HasValue)
        {
            options.Dpi = RenderDpi.Value;
        }

        if (MinimumConfidence.HasValue)
        {
            options.MinimumConfidence = MinimumConfidence.Value;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        PdfDocument document = PdfDocument.Load(inputPath);
        PdfSearchableOcrResult result = await document
            .MakeSearchableAsync(session.Engine, options, CancelToken)
            .ConfigureAwait(false);
        await result.Document.SaveAsync(
                outputPath,
                Force.IsPresent
                    ? OfficeConversionFileConflictPolicy.Replace
                    : OfficeConversionFileConflictPolicy.FailIfExists,
                CancelToken)
            .ConfigureAwait(false);
        WriteObject(PassThru.IsPresent ? result : new FileInfo(outputPath));
    }
}
