using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Gets or extracts image resources from a PDF.</summary>
/// <example>
///   <summary>Extract images from selected pages.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$proof = @(
///     Get-OfficePdfImage -Path .\Examples\Documents\Report.pdf -PageRange '1-2'
///     Get-OfficePdfImage -Path .\Examples\Documents\Report.pdf -OutputDirectory .\Examples\Documents\PdfImages -BaseName 'report-image'
/// )
/// $proof</code>
///   <para>Returns image metadata or writes extracted images to disk.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfImage")]
[OutputType(typeof(PdfExtractedImage), typeof(FileInfo))]
public sealed class GetOfficePdfImageCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional page ranges such as 1-3,5.</summary>
    [Parameter]
    public string? PageRange { get; set; }

    /// <summary>Optional directory where images should be written.</summary>
    [Parameter]
    public string? OutputDirectory { get; set; }

    /// <summary>Base file name used when extracting images to disk.</summary>
    [Parameter]
    public string BaseName { get; set; } = "image";

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var inputPath = PdfCommandUtilities.ResolvePath(this, Path);
        var document = PdfDocument.Open(inputPath);
        if (!string.IsNullOrWhiteSpace(OutputDirectory))
        {
            var outputDirectory = PdfCommandUtilities.ResolvePath(this, OutputDirectory!);
            PdfCommandUtilities.EnsureOutputDirectory(outputDirectory);
            var paths = string.IsNullOrWhiteSpace(PageRange)
                ? document.Read.SaveImages(outputDirectory, BaseName)
                : document.Read.SaveImages(outputDirectory, PdfPageSelection.Parse(PageRange!), BaseName);
            foreach (var path in paths)
            {
                WriteObject(new FileInfo(path));
            }

            return;
        }

        var images = string.IsNullOrWhiteSpace(PageRange)
            ? document.Read.Images()
            : document.Read.Images(PageRange!);
        foreach (var image in images)
        {
            WriteObject(image);
        }
    }
}
