using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Runs an external OCR provider and merges recognized words with native PDF text.</summary>
[Cmdlet(VerbsLifecycle.Invoke, "OfficePdfOcrMerge")]
[OutputType(typeof(PdfOcrMergeResult))]
public sealed class InvokeOfficePdfOcrMergeCommand : AsyncPSCmdlet
{
    /// <summary>Source PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>External OCR provider implementation.</summary>
    [Parameter(Mandatory = true)]
    public IPdfOcrProvider Provider { get; set; } = null!;

    /// <summary>Optional page selection, DPI, confidence, overlap, and limits.</summary>
    [Parameter]
    public PdfOcrMergeOptions? Options { get; set; }

    /// <summary>Optional bounded PDF parsing settings.</summary>
    [Parameter]
    public PdfReadOptions? ReadOptions { get; set; }

    /// <summary>Password used to authenticate an encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed extraction restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync()
    {
        var readOptions = PdfCommandUtilities.CreateReadOptions(
            ReadOptions,
            Password,
            IgnorePermissionRestrictions.IsPresent);
        var document = PdfCommandUtilities.LoadDocument(SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path), readOptions);
        WriteObject(await document.Read.OcrAsync(Provider, Options, readOptions, CancelToken).ConfigureAwait(false));
    }
}
