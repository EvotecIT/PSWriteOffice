using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Imports bounded DTD-free XFDF through the validated PDF form filler.</summary>
[Cmdlet(VerbsData.Import, "OfficePdfXfdf", DefaultParameterSetName = "Text", SupportsShouldProcess = true)]
[OutputType(typeof(PdfDocument))]
public sealed class ImportOfficePdfXfdfCommand : PSCmdlet
{
    /// <summary>Source PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>XFDF XML.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Text")]
    public string Xfdf { get; set; } = string.Empty;

    /// <summary>Path to an XFDF file.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "File")]
    public string XfdfPath { get; set; } = string.Empty;

    /// <summary>Destination PDF path.</summary>
    [Parameter(Mandatory = true)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Optional validated form filling behavior.</summary>
    [Parameter]
    public PdfFormFillerOptions? Options { get; set; }

    /// <summary>Return the rewritten fluent PDF document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        if (!ShouldProcess(output, "Import XFDF into PDF form fields")) return;
        var xml = ParameterSetName == "File"
            ? File.ReadAllText(SessionState.Path.GetUnresolvedProviderPathFromPSPath(XfdfPath))
            : Xfdf;
        var result = PdfDocument.Load(input).Forms.ImportXfdf(xml, Options);
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        result.Save(output);
        if (PassThru.IsPresent) WriteObject(result);
    }
}
