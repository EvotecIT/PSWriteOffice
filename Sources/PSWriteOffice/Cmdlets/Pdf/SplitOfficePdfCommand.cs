using System.Globalization;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Splits a PDF into one file per page.</summary>
/// <example>
///   <summary>Split a PDF into page files.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$pages = Split-OfficePdf -Path .\Examples\Documents\Combined.pdf -OutputDirectory .\Examples\Documents\Pages -Prefix 'combined-page'
/// $pages | Select-Object Name, Length</code>
///   <para>Creates one output PDF for each page and returns the written files.</para>
/// </example>
[Cmdlet(VerbsCommon.Split, "OfficePdf")]
[OutputType(typeof(FileInfo))]
public sealed class SplitOfficePdfCommand : PSCmdlet
{
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Output directory.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputDirectory { get; set; } = string.Empty;

    /// <summary>Output file prefix.</summary>
    [Parameter]
    public string Prefix { get; set; } = "page";

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var outputDirectory = PdfCommandUtilities.ResolvePath(this, OutputDirectory);
        Directory.CreateDirectory(outputDirectory);
        var documents = PdfDocument.Open(PdfCommandUtilities.ResolvePath(this, Path)).Pages.Split();
        for (var i = 0; i < documents.Count; i++)
        {
            var outputPath = System.IO.Path.Combine(outputDirectory, Prefix + "-" + (i + 1).ToString(CultureInfo.InvariantCulture) + ".pdf");
            documents[i].Save(outputPath);
            WriteObject(new FileInfo(outputPath));
        }
    }
}
