using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Applies rectangle-based PDF redactions and writes a new PDF.</summary>
/// <example>
///   <summary>Apply a redaction rectangle.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficePdfRedacted -Path .\Report.pdf -OutputPath .\Report-Redacted.pdf -PageNumber 1 -X 72 -Y 650 -Width 240 -Height 32</code>
///   <para>Removes matching text objects and annotations in the rectangle, then paints a redaction mark.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficePdfRedacted", DefaultParameterSetName = ParameterSetRectangle, SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class ConvertToOfficePdfRedactedCommand : PSCmdlet
{
    private const string ParameterSetRectangle = "Rectangle";
    private const string ParameterSetArea = "Area";

    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>One-based page number for the redaction rectangle.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetRectangle)]
    public int PageNumber { get; set; }

    /// <summary>Left coordinate in PDF points.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetRectangle)]
    public double X { get; set; }

    /// <summary>Bottom coordinate in PDF points.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetRectangle)]
    public double Y { get; set; }

    /// <summary>Rectangle width in PDF points.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetRectangle)]
    public double Width { get; set; }

    /// <summary>Rectangle height in PDF points.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetRectangle)]
    public double Height { get; set; }

    /// <summary>Optional redaction area label.</summary>
    [Parameter(ParameterSetName = ParameterSetRectangle)]
    public string? Label { get; set; }

    /// <summary>One or more pre-created OfficeIMO.Pdf redaction areas.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetArea)]
    public PdfRedactionArea[] Area { get; set; } = System.Array.Empty<PdfRedactionArea>();

    /// <summary>Redaction fill color in #RRGGBB format. Defaults to black.</summary>
    [Parameter]
    public string? FillColor { get; set; }

    /// <summary>Paint only areas that match text or annotations in the redaction plan.</summary>
    [Parameter]
    public SwitchParameter OnlyPaintMatches { get; set; }

    /// <summary>Password used to read a Standard password-encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        string inputPath = PdfCommandUtilities.ResolvePath(this, Path);
        string outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!ShouldProcess(outputPath, "Write redacted PDF"))
        {
            return;
        }

        PdfRedactionArea[] areas = ParameterSetName == ParameterSetArea
            ? Area
            : new[] { new PdfRedactionArea(PageNumber, X, Y, Width, Height, Label) };

        var options = new PdfRedactionApplyOptions
        {
            FillColor = PdfCommandUtilities.ParseColor(FillColor) ?? PdfColor.Black,
            PaintUnmatchedAreas = !OnlyPaintMatches.IsPresent
        };

        PdfCommandUtilities.EnsureDirectory(outputPath);
        PdfRedactionApplier.Apply(inputPath, outputPath, areas, options, readOptions: PdfCommandUtilities.CreateReadOptions(Password));
        WriteObject(new FileInfo(outputPath));
    }
}
