using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets generated PDF compliance profile and readiness groundwork.</summary>
[Cmdlet(VerbsCommon.Set, "OfficePdfCompliance", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfCompliance")]
[OutputType(typeof(PdfDocument))]
public sealed class SetOfficePdfComplianceCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Requested generated PDF compliance profile.</summary>
    [Parameter(Mandatory = true)]
    public PdfComplianceProfile Profile { get; set; }

    /// <summary>Configure common PDF/A or PDF/UA groundwork for the selected profile.</summary>
    [Parameter]
    public SwitchParameter Groundwork { get; set; }

    /// <summary>Catalog language used by compliance groundwork.</summary>
    [Parameter]
    public string Language { get; set; } = "en-US";

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ResolveDocument(this, Document, ParameterSetName, ParameterSetDocument);
        if (Groundwork.IsPresent)
        {
            ApplyGroundwork(document);
        }

        document.Compliance(Profile);
        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }

    private void ApplyGroundwork(PdfDocument document)
    {
        switch (Profile)
        {
            case PdfComplianceProfile.PdfUa1:
                document.ConfigurePdfUaGroundwork(Language);
                break;
            case PdfComplianceProfile.PdfA2B:
            case PdfComplianceProfile.PdfA2U:
            case PdfComplianceProfile.PdfA2A:
            case PdfComplianceProfile.PdfA3B:
            case PdfComplianceProfile.PdfA3U:
            case PdfComplianceProfile.PdfA3A:
                document.ConfigurePdfAGroundwork(Profile, Language);
                break;
            default:
                WriteWarning("Groundwork is currently available for PDF/A and PDF/UA profiles.");
                break;
        }
    }
}
