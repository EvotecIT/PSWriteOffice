using System.IO;
using System.Management.Automation;
using OfficeIMO.Latex;

namespace PSWriteOffice.Cmdlets.Latex;

/// <summary>Saves an OfficeIMO LaTeX document.</summary>
[Cmdlet(VerbsData.Save, "OfficeLatex", SupportsShouldProcess = true)]
[OutputType(typeof(LatexDocument))]
public sealed class SaveOfficeLatexCommand : PSCmdlet
{
    /// <summary>LaTeX document to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public LatexDocument Document { get; set; } = null!;

    /// <summary>Destination path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional writer settings.</summary>
    [Parameter]
    public LatexWriterOptions? Options { get; set; }

    /// <summary>Return the saved document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var path = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        if (!ShouldProcess(path, "Save LaTeX document")) return;
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        Document.Save(path, Options);
        if (PassThru.IsPresent) WriteObject(Document);
    }
}
