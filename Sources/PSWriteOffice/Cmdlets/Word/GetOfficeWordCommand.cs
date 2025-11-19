using System.IO;
using System.Management.Automation;
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
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath", "Path")]
    public string InputPath { get; set; } = string.Empty;

    [Parameter]
    public SwitchParameter ReadOnly { get; set; }

    [Parameter]
    public SwitchParameter AutoSave { get; set; }

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
