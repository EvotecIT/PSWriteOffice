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

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var options = Options ?? new PdfRewritePreservationOptions();
        var original = PdfCommandUtilities.LoadDocument(
            SessionState.Path.GetUnresolvedProviderPathFromPSPath(ReferencePath), options.OriginalReadOptions);
        var rewritten = PdfCommandUtilities.LoadDocument(
            SessionState.Path.GetUnresolvedProviderPathFromPSPath(DifferencePath), options.RewrittenReadOptions);
        var report = original.AssessRewritePreservation(rewritten, options);
        if (FailOnLoss.IsPresent) report.ThrowIfFailed();
        WriteObject(report);
    }
}
