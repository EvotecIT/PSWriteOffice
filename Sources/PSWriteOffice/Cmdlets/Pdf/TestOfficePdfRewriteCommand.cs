using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates a proof report for user-visible signals preserved by a PDF rewrite.</summary>
[Cmdlet(VerbsDiagnostic.Test, "OfficePdfRewrite")]
[OutputType(typeof(PdfRewritePreservationReport))]
public sealed class TestOfficePdfRewriteCommand : PSCmdlet
{
    /// <summary>Original PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string ReferencePath { get; set; } = string.Empty;

    /// <summary>Rewritten PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string DifferencePath { get; set; } = string.Empty;

    /// <summary>Optional required preservation signals and limits.</summary>
    [Parameter]
    public PdfRewritePreservationOptions? Options { get; set; }

    /// <summary>Throw when preservation checks find a mismatch.</summary>
    [Parameter]
    public SwitchParameter FailOnLoss { get; set; }

    /// <summary>Password used to authenticate the original PDF.</summary>
    [Parameter]
    public string? ReferencePassword { get; set; }

    /// <summary>After authentication, explicitly ignore restrictions on the original PDF.</summary>
    [Parameter]
    public SwitchParameter IgnoreReferencePermissionRestrictions { get; set; }

    /// <summary>Password used to authenticate the rewritten PDF.</summary>
    [Parameter]
    public string? DifferencePassword { get; set; }

    /// <summary>After authentication, explicitly ignore restrictions on the rewritten PDF.</summary>
    [Parameter]
    public SwitchParameter IgnoreDifferencePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var options = Options ?? new PdfRewritePreservationOptions();
        var original = PdfCommandUtilities.LoadDocument(
            SessionState.Path.GetUnresolvedProviderPathFromPSPath(ReferencePath),
            PdfCommandUtilities.CreateReadOptions(
                options.OriginalReadOptions,
                ReferencePassword,
                IgnoreReferencePermissionRestrictions.IsPresent));
        var rewritten = PdfCommandUtilities.LoadDocument(
            SessionState.Path.GetUnresolvedProviderPathFromPSPath(DifferencePath),
            PdfCommandUtilities.CreateReadOptions(
                options.RewrittenReadOptions,
                DifferencePassword,
                IgnoreDifferencePermissionRestrictions.IsPresent));
        var report = original.AssessRewritePreservation(rewritten, options);
        if (FailOnLoss.IsPresent) report.ThrowIfFailed();
        WriteObject(report);
    }
}
