using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a page break to a PDF document.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePdfPageBreak", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfPageBreak")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfPageBreakCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ResolveDocument(this, Document, ParameterSetName, ParameterSetDocument);
        document.PageBreak();
        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
