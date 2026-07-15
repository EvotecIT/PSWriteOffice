using System.IO;
using System.Management.Automation;
using OfficeIMO.AsciiDoc;

namespace PSWriteOffice.Cmdlets.AsciiDoc;

/// <summary>Saves an OfficeIMO AsciiDoc document.</summary>
[Cmdlet(VerbsData.Save, "OfficeAsciiDoc", SupportsShouldProcess = true)]
[OutputType(typeof(AsciiDocDocument))]
public sealed class SaveOfficeAsciiDocCommand : PSCmdlet
{
    /// <summary>AsciiDoc document to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public AsciiDocDocument Document { get; set; } = null!;

    /// <summary>Destination path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional writer settings.</summary>
    [Parameter]
    public AsciiDocWriterOptions? Options { get; set; }

    /// <summary>Return the saved document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var path = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        if (!ShouldProcess(path, "Save AsciiDoc document")) return;
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        Document.Save(path, Options);
        if (PassThru.IsPresent) WriteObject(Document);
    }
}
