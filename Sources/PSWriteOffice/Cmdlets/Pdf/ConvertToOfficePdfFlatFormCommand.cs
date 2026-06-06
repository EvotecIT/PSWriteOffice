using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Converts a PDF with simple AcroForm fields into a flat PDF.</summary>
/// <example>
///   <summary>Flatten a filled PDF form.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Set-OfficePdfForm -Path .\Examples\Documents\Request.pdf -OutputPath .\Examples\Documents\Request-Filled.pdf -Field @{
///     Requester = 'Ada Lovelace'
///     Priority = 'High'
/// }
/// ConvertTo-OfficePdfFlatForm -Path .\Examples\Documents\Request-Filled.pdf -OutputPath .\Examples\Documents\Request-Flat.pdf</code>
///   <para>Turns simple form fields into static page content.</para>
/// </example>
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
