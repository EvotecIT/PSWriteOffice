using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Opens an existing Word document.</summary>
/// <para>Returns an OfficeIMO <see cref="WordDocument"/> for inspection or advanced operations.</para>
/// <example>
///   <summary>Load a document in read-only mode.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx -ReadOnly</code>
///   <para>Loads <c>Report.docx</c> and exposes the document object for querying.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWord")]
public sealed class GetOfficeWordCommand : PSCmdlet
{
    /// <summary>Path to the .docx. Accepts PS paths.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath", "Path")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Open in read-only mode.</summary>
    [Parameter]
    public SwitchParameter ReadOnly { get; set; }

    /// <summary>Enable AutoSave when editing.</summary>
    [Parameter]
    public SwitchParameter AutoSave { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var fullPath = ResolvePath();
        var document = WordDocumentService.LoadDocument(fullPath, ReadOnly.IsPresent, AutoSave.IsPresent);
        WriteObject(document);
    }

    private string ResolvePath()
    {
        var providerPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
        return Path.IsPathRooted(providerPath)
            ? providerPath
            : Path.Combine(SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }
}
