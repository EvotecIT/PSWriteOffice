using System.Management.Automation;
using OfficeIMO.Email;

namespace PSWriteOffice.Cmdlets.Email;

/// <summary>Reads a native EML, MSG, or TNEF artifact with bounded diagnostics.</summary>
[Cmdlet(VerbsCommon.Get, "OfficeEmail")]
[OutputType(typeof(EmailDocument), typeof(EmailReadResult))]
public sealed class GetOfficeEmailCommand : PSCmdlet
{
    /// <summary>Path to an EML, MSG, TNEF, or winmail.dat file.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional format detection, compound-file, MIME, attachment, and size limits.</summary>
    [Parameter]
    public EmailReaderOptions? Options { get; set; }

    /// <summary>Return the read result with diagnostics and consumed byte count.</summary>
    [Parameter]
    public SwitchParameter AsResult { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var result = new EmailDocumentReader(Options ?? EmailReaderOptions.Default).Read(
            SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path));
        WriteObject(AsResult.IsPresent ? result : result.Document);
    }
}
