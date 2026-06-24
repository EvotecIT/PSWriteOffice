using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets generated PDF compliance profile and readiness groundwork.</summary>
/// <example>
///   <summary>Configure PDF/A groundwork before saving.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfCompliance.pdf {
///     Set-OfficePdfCompliance -Profile PdfA3B -Groundwork -Language 'en-US'
///     Add-OfficePdfHeading -Text 'Compliance-ready report'
///     Get-OfficePdfCompliance -Profile PdfA3B
/// }</code>
///   <para>Applies OfficeIMO.Pdf compliance groundwork and emits a readiness report inside the DSL.</para>
/// </example>
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
        else
        {
            document.Compliance(Profile);
        }
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
            case PdfComplianceProfile.PdfUa2:
                document.ConfigurePdfUaGroundwork(Profile, Language);
                break;
            case PdfComplianceProfile.PdfA2B:
            case PdfComplianceProfile.PdfA2U:
            case PdfComplianceProfile.PdfA2A:
            case PdfComplianceProfile.PdfA3B:
            case PdfComplianceProfile.PdfA3U:
            case PdfComplianceProfile.PdfA3A:
            case PdfComplianceProfile.PdfA4:
            case PdfComplianceProfile.PdfA4E:
            case PdfComplianceProfile.PdfA4F:
                document.ConfigurePdfAGroundwork(Profile, Language);
                break;
            default:
                WriteWarning("Groundwork is currently available for PDF/A and PDF/UA profiles.");
                break;
        }
    }
}
