using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets or clears the generated PDF page border decoration.</summary>
[Cmdlet(VerbsCommon.Set, "OfficePdfPageBorder", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfPageBorder")]
[OutputType(typeof(PdfDocument))]
public sealed class SetOfficePdfPageBorderCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Border color in #RRGGBB format.</summary>
    [Parameter]
    public string? Color { get; set; }

    /// <summary>Border stroke width in PDF points.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Distance from the page edge to the border path in PDF points.</summary>
    [Parameter]
    public double? Inset { get; set; }

    /// <summary>Border opacity from 0 through 1.</summary>
    [Parameter]
    public double? Opacity { get; set; }

    /// <summary>Border dash style.</summary>
    [Parameter]
    public OfficeStrokeDashStyle DashStyle { get; set; } = OfficeStrokeDashStyle.Solid;

    /// <summary>Clear the generated PDF page border decoration.</summary>
    [Parameter]
    public SwitchParameter Clear { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ResolveDocument(this, Document, ParameterSetName, ParameterSetDocument);
        if (Clear.IsPresent)
        {
            document.PageBorder(null);
        }
        else
        {
            document.PageBorder(PdfCommandUtilities.ParseColor(Color), Width, Inset, Opacity, DashStyle);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
