using System.IO;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Exports a Word page as PNG or SVG with structured image diagnostics.</summary>
/// <example>
///   <summary>Export the first page as SVG.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Export-OfficeWordImage -Path .\Report.docx -OutputPath .\Report.svg -Format Svg</code>
///   <para>Returns the OfficeIMO image export result after writing the image.</para>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeWordImage", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeImageExportResult))]
public sealed class ExportOfficeWordImageCommand : PSCmdlet
{
    /// <summary>Path to the Word document.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open Word document instance.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document")]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Destination PNG or SVG path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Output image format.</summary>
    [Parameter]
    public OfficeImageExportFormat Format { get; set; } = OfficeImageExportFormat.Png;

    /// <summary>Optional page, size, scale, theme, and rendering settings.</summary>
    [Parameter]
    public WordImageExportOptions? Options { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        if (!ShouldProcess(output, $"Export Word page as {Format}")) return;
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        WordDocument? owned = null;
        try
        {
            var document = Document;
            if (ParameterSetName == "Path")
            {
                var input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
                owned = WordDocumentService.LoadDocument(input, readOnly: true, autoSave: false);
                document = owned;
            }
            WriteObject(Format == OfficeImageExportFormat.Svg
                ? document.SaveAsSvg(output, Options)
                : document.SaveAsPng(output, Options));
        }
        finally
        {
            owned?.Dispose();
        }
    }
}
