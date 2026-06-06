using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets or clears the generated PDF page background color.</summary>
[Cmdlet(VerbsCommon.Set, "OfficePdfBackground", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfBackground")]
[OutputType(typeof(PdfDocument))]
public sealed class SetOfficePdfBackgroundCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Background color in #RRGGBB format.</summary>
    [Parameter]
    public string? Color { get; set; }

    /// <summary>Clear the generated PDF page background color.</summary>
    [Parameter]
    public SwitchParameter Clear { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ResolveDocument(this, Document, ParameterSetName, ParameterSetDocument);
        document.Background(Clear.IsPresent ? null : PdfCommandUtilities.ParseColor(Color));
        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
