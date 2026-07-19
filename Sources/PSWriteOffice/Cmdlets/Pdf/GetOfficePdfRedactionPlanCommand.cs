using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Previews text and annotations intersecting rectangle-based redaction areas.</summary>
/// <remarks>
/// This command reports redaction impact only. It does not remove or rewrite PDF content.
/// </remarks>
/// <example>
///   <summary>Preview a redaction rectangle.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePdfRedactionPlan -Path .\Report.pdf -PageNumber 1 -X 72 -Y 650 -Width 240 -Height 32</code>
///   <para>Returns line-level text blocks and annotations that intersect the rectangle.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfRedactionPlan", DefaultParameterSetName = ParameterSetRectangle)]
[OutputType(typeof(PdfRedactionPlan))]
public sealed class GetOfficePdfRedactionPlanCommand : PSCmdlet
{
    private const string ParameterSetRectangle = "Rectangle";
    private const string ParameterSetArea = "Area";

    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

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

    /// <summary>Password used to read a Standard password-encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var areas = ParameterSetName == ParameterSetArea
            ? Area
            : new[] { new PdfRedactionArea(PageNumber, X, Y, Width, Height, Label) };

        WriteObject(PdfDocument
            .Open(PdfCommandUtilities.ResolvePath(this, Path), PdfCommandUtilities.CreateReadOptions(Password))
            .PlanRedactions(areas));
    }
}
