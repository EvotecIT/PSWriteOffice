using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets PDF page size, orientation, and margins.</summary>
[Cmdlet(VerbsCommon.Set, "OfficePdfPageSetup", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfPageSetup")]
[OutputType(typeof(PdfDocument))]
public sealed class SetOfficePdfPageSetupCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Page size name: A4, A5, Letter, Legal, or Custom.</summary>
    [Parameter]
    public string? PageSize { get; set; }

    /// <summary>Custom page width in PDF points when -PageSize Custom is used.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Custom page height in PDF points when -PageSize Custom is used.</summary>
    [Parameter]
    public double? Height { get; set; }

    /// <summary>Use landscape orientation.</summary>
    [Parameter]
    public SwitchParameter Landscape { get; set; }

    /// <summary>Uniform margin in PDF points.</summary>
    [Parameter]
    public double? Margin { get; set; }

    /// <summary>Left margin in PDF points.</summary>
    [Parameter]
    public double? Left { get; set; }

    /// <summary>Top margin in PDF points.</summary>
    [Parameter]
    public double? Top { get; set; }

    /// <summary>Right margin in PDF points.</summary>
    [Parameter]
    public double? Right { get; set; }

    /// <summary>Bottom margin in PDF points.</summary>
    [Parameter]
    public double? Bottom { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ResolveDocument(this, Document, ParameterSetName, ParameterSetDocument);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(PageSize)) ||
            MyInvocation.BoundParameters.ContainsKey(nameof(Width)) ||
            MyInvocation.BoundParameters.ContainsKey(nameof(Height)) ||
            MyInvocation.BoundParameters.ContainsKey(nameof(Landscape)))
        {
            document.Size(PdfCommandUtilities.ResolvePageSize(PageSize, Width, Height, Landscape.IsPresent));
        }

        if (Margin.HasValue)
        {
            document.Margin(Margin.Value);
        }
        else if (Left.HasValue || Top.HasValue || Right.HasValue || Bottom.HasValue)
        {
            document.Margin(Left ?? 72, Top ?? 72, Right ?? 72, Bottom ?? 72);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
