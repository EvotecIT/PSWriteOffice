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

    /// <summary>Writer mode. Preserve retains unchanged source; Canonical normalizes output.</summary>
    [Parameter]
    public LatexWriterMode? Mode { get; set; }

    /// <summary>Canonical line ending: LF, CRLF, or CR. Omit it to retain the source preference.</summary>
    [Parameter]
    [ValidateSet("LF", "CRLF", "CR")]
    public string? LineEnding { get; set; }

    /// <summary>Return the saved document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var path = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        if (!ShouldProcess(path, "Save LaTeX document")) return;
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        Document.Save(path, BuildOptions());
        if (PassThru.IsPresent) WriteObject(Document);
    }

    private LatexWriterOptions BuildOptions() {
        var options = new LatexWriterOptions {
            Mode = Options?.Mode ?? LatexWriterMode.Preserve,
            LineEnding = Options?.LineEnding
        };
        if (Mode.HasValue) options.Mode = Mode.Value;
        if (LineEnding != null) options.LineEnding = ResolveLineEnding(LineEnding);
        return options;
    }

    private static string ResolveLineEnding(string value) => value switch {
        "LF" => "\n",
        "CRLF" => "\r\n",
        "CR" => "\r",
        _ => throw new PSArgumentException("LineEnding must be LF, CRLF, or CR.", nameof(LineEnding))
    };
}
