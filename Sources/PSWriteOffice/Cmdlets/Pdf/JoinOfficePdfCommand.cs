using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Joins multiple PDF files into a single PDF.</summary>
/// <example>
///   <summary>Join two PDFs in order.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$cover = '.\Examples\Documents\Cover.pdf'
/// $report = '.\Examples\Documents\Report.pdf'
/// Join-OfficePdf -Path $cover, $report -OutputPath .\Examples\Documents\Combined.pdf -PassThru
/// Get-OfficePdfInfo -Path .\Examples\Documents\Combined.pdf | Select-Object PageCount</code>
///   <para>Writes a single PDF containing the input documents in the requested order, then checks the result.</para>
/// </example>
[Cmdlet(VerbsCommon.Join, "OfficePdf", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class JoinOfficePdfCommand : PSCmdlet
{
    /// <summary>Input PDF paths in output order.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath")]
    public string[] Path { get; set; } = System.Array.Empty<string>();

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Emit the saved file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Flatten visual annotation appearances before merging.</summary>
    [Parameter]
    public SwitchParameter FlattenVisualAnnotations { get; set; }

    /// <summary>Resize each merged page to a known OfficeIMO page size such as A4, Letter, or Custom.</summary>
    [Parameter]
    public string? PageSize { get; set; }

    /// <summary>Custom output page width in points when -PageSize Custom is used.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Custom output page height in points when -PageSize Custom is used.</summary>
    [Parameter]
    public double? Height { get; set; }

    /// <summary>Use the landscape orientation of the selected output page size.</summary>
    [Parameter]
    public SwitchParameter Landscape { get; set; }

    /// <summary>How source page content is fitted into the resized output page.</summary>
    [Parameter]
    public PdfPageResizeMode ResizeMode { get; set; } = PdfPageResizeMode.Fit;

    /// <summary>Margin, in points, reserved around resized page content.</summary>
    [Parameter]
    public double? ResizeMargin { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write joined PDF"))
        {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        var resizeOptions = PdfCommandUtilities.CreatePageResizeOptions(
            PageSize,
            Width,
            Height,
            Landscape.IsPresent,
            ResizeMode,
            ResizeMargin,
            MyInvocation.BoundParameters.ContainsKey(nameof(PageSize)) ||
            MyInvocation.BoundParameters.ContainsKey(nameof(Width)) ||
            MyInvocation.BoundParameters.ContainsKey(nameof(Height)) ||
            Landscape.IsPresent ||
            MyInvocation.BoundParameters.ContainsKey(nameof(ResizeMode)) ||
            ResizeMargin.HasValue);

        var documents = Path
            .Select(path => PdfDocument.Open(PdfCommandUtilities.ResolvePath(this, path)))
            .ToArray();
        if (documents.Length == 0)
        {
            throw new PSArgumentException("Provide at least one input PDF path.", nameof(Path));
        }

        PdfDocument Prepare(PdfDocument document)
        {
            var prepared = FlattenVisualAnnotations.IsPresent
                ? document.FlattenVisualAnnotations()
                : document;
            return resizeOptions == null
                ? prepared
                : prepared.Pages.Resize(resizeOptions);
        }

        var combined = PdfDocument.Merge(documents.Select(Prepare));

        combined.Save(outputPath).RequireSuccess();

        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(outputPath));
        }
    }
}
