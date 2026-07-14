using System.IO;
using System.Management.Automation;
using OfficeIMO.AsciiDoc;
using OfficeIMO.AsciiDoc.Markdown;

namespace PSWriteOffice.Cmdlets.AsciiDoc;

/// <summary>Converts AsciiDoc to Markdown with fidelity diagnostics.</summary>
[Cmdlet(VerbsData.ConvertTo, "OfficeAsciiDocMarkdown", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true)]
[OutputType(typeof(AsciiDocToMarkdownResult))]
public sealed class ConvertToOfficeAsciiDocMarkdownCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to an AsciiDoc file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    public string Path { get; set; } = string.Empty;

    /// <summary>AsciiDoc document to convert.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public AsciiDocDocument Document { get; set; } = null!;

    /// <summary>Optional Markdown destination path.</summary>
    [Parameter]
    public string? OutputPath { get; set; }

    /// <summary>Optional conversion settings.</summary>
    [Parameter]
    public AsciiDocToMarkdownOptions? Options { get; set; }

    /// <summary>Throw when a source feature cannot be mapped exactly.</summary>
    [Parameter]
    public SwitchParameter FailOnLoss { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = ParameterSetName == ParameterSetPath
            ? AsciiDocDocument.Load(SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path)).Document
            : Document;
        var result = document.ToMarkdownDocumentResult(Options);
        if (FailOnLoss.IsPresent) result.RequireNoLoss();
        if (!string.IsNullOrWhiteSpace(OutputPath))
        {
            var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
            if (!ShouldProcess(output, "Write converted Markdown")) return;
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
            File.WriteAllText(output, result.Value.ToMarkdown());
        }
        WriteObject(result);
    }
}
