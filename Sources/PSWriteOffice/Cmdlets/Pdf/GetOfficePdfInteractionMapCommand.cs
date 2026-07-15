using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Builds text-selection and interactive hit regions for one PDF page.</summary>
[Cmdlet(VerbsCommon.Get, "OfficePdfInteractionMap")]
[OutputType(typeof(PdfPageInteractionMap))]
public sealed class GetOfficePdfInteractionMapCommand : PSCmdlet
{
    /// <summary>Source PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>One-based page number.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int Page { get; set; } = 1;

    /// <summary>Optional text-region limits.</summary>
    [Parameter]
    public PdfPageInteractionOptions? Options { get; set; }

    /// <summary>Optional bounded PDF parsing settings.</summary>
    [Parameter]
    public PdfReadOptions? ReadOptions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() => WriteObject(PdfCommandUtilities.LoadDocument(
        SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path), ReadOptions).Read.Interactions(Page, Options, ReadOptions));
}
