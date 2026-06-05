using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Converts a PDF with simple AcroForm fields into a flat PDF.</summary>
[Cmdlet(VerbsData.ConvertTo, "OfficePdfFlatForm")]
[OutputType(typeof(FileInfo))]
public sealed class ConvertToOfficePdfFlatFormCommand : PSCmdlet
{
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var result = PdfDocument.Open(PdfCommandUtilities.ResolvePath(this, Path)).Forms.Flatten();
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        PdfCommandUtilities.EnsureDirectory(outputPath);
        result.Save(outputPath);
        WriteObject(new FileInfo(outputPath));
    }
}
