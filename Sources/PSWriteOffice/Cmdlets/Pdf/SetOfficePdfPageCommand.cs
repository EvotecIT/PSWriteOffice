using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets page-level PDF properties and writes a new PDF.</summary>
/// <example>
///   <summary>Rotate selected PDF pages.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$proof = @(
///     Set-OfficePdfPage -Path .\Examples\Documents\Scanned.pdf -PageRange '2,4' -Rotation 90 -OutputPath .\Examples\Documents\Scanned-Rotated.pdf
///     Get-OfficePdfInfo -Path .\Examples\Documents\Scanned-Rotated.pdf | Select-Object PageCount
/// )
/// $proof</code>
///   <para>Rotates selected pages and writes a new PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfPage")]
[OutputType(typeof(FileInfo))]
public sealed class SetOfficePdfPageCommand : PSCmdlet
{
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Page ranges such as 1-3,5. Omit to affect all pages.</summary>
    [Parameter]
    public string? PageRange { get; set; }

    /// <summary>Rotation in degrees. Supported values are 0, 90, 180, and 270.</summary>
    [Parameter(Mandatory = true)]
    [ValidateSet("0", "90", "180", "270")]
    public int Rotation { get; set; }

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true)]
    public string OutputPath { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        PdfCommandUtilities.EnsureDirectory(outputPath);
        var document = PdfDocument.Open(PdfCommandUtilities.ResolvePath(this, Path));
        var result = string.IsNullOrWhiteSpace(PageRange)
            ? document.Pages.Rotate(Rotation)
            : document.Pages.Rotate(Rotation, PageRange!);
        result.Save(outputPath);
        WriteObject(new FileInfo(outputPath));
    }
}
