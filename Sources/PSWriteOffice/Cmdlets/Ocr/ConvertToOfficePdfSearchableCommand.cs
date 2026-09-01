using System.IO;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using OfficeIMO.Reader.Ocr;
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
///   <code>ConvertTo-OfficePdfSearchable -Path .\Scan.pdf -OutputPath .\Searchable.pdf -Language eng+pol -PassThru</code>
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

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync()
    {
        string inputPath = PdfCommandUtilities.ResolveExistingFilePath(this, Path);
        string outputPath = PdfCommandUtilities.ResolveOutputFilePath(this, OutputPath, ".pdf", Force.IsPresent);
        if (!ShouldProcess(outputPath, "Create searchable PDF"))
        {
            return;
        }

        OfficeOcrOptions options = CreateOptions();
        if (RenderDpi.HasValue)
        {
            options.Pdf.Dpi = RenderDpi.Value;
        }

        if (MinimumConfidence.HasValue)
        {
            options.Pdf.MinimumConfidence = MinimumConfidence.Value;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        PdfSearchableOcrResult result = await OfficeOcr
            .MakePdfSearchableAsync(inputPath, outputPath, options, CancelToken)
            .ConfigureAwait(false);
        WriteObject(PassThru.IsPresent ? result : new FileInfo(outputPath));
    }
}
