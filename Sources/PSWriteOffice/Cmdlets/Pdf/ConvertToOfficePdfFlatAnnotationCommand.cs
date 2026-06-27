using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Flattens supported visual PDF annotations into static page content.</summary>
/// <example>
///   <summary>Flatten visual annotations.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficePdfFlatAnnotation -Path .\Reviewed.pdf -OutputPath .\Reviewed-Flat.pdf</code>
///   <para>Paints supported annotation appearance streams into page content and writes a new PDF.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficePdfFlatAnnotation", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class ConvertToOfficePdfFlatAnnotationCommand : PSCmdlet
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
        var result = PdfDocument.Open(PdfCommandUtilities.ResolvePath(this, Path)).FlattenVisualAnnotations();
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write flattened annotation PDF"))
        {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        result.Save(outputPath);
        WriteObject(new FileInfo(outputPath));
    }
}
