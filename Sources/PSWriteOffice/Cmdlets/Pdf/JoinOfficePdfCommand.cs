using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Joins multiple PDF files into a single PDF.</summary>
/// <example>
///   <summary>Join two PDFs in order.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Join-OfficePdf -Path .\Cover.pdf, .\Report.pdf -OutputPath .\Combined.pdf -PassThru</code>
///   <para>Writes a single PDF containing the input documents in the requested order.</para>
/// </example>
[Cmdlet(VerbsCommon.Join, "OfficePdf")]
[OutputType(typeof(FileInfo))]
public sealed class JoinOfficePdfCommand : PSCmdlet
{
    /// <summary>Input PDF paths in output order.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath")]
    public string[] Path { get; set; } = System.Array.Empty<string>();

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Emit the saved file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        PdfCommandUtilities.EnsureDirectory(outputPath);
        PdfMerger.MergeFiles(Path.Select(path => PdfCommandUtilities.ResolvePath(this, path)), outputPath);
        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(outputPath));
        }
    }
}
