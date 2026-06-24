using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Removes PDF annotations matching friendly filters.</summary>
[Cmdlet(VerbsCommon.Remove, "OfficePdfAnnotation", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
[OutputType(typeof(PdfAnnotationEditResult))]
public sealed class RemoveOfficePdfAnnotationCommand : PSCmdlet
{
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Specific annotation object number to remove.</summary>
    [Parameter]
    public int? ObjectNumber { get; set; }

    /// <summary>Optional one-based page number filter.</summary>
    [Parameter]
    public int? PageNumber { get; set; }

    /// <summary>Optional annotation subtype filter such as Text, Link, Widget, or FreeText.</summary>
    [Parameter]
    public string? Subtype { get; set; }

    /// <summary>Keep popup annotations linked from matching annotations through /Popup.</summary>
    [Parameter]
    public SwitchParameter KeepPopups { get; set; }

    /// <summary>Return the annotation edit result instead of the output file.</summary>
    [Parameter]
    public SwitchParameter PassThruReport { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        string inputPath = PdfCommandUtilities.ResolvePath(this, Path);
        string outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!ShouldProcess(outputPath, "Remove PDF annotations"))
        {
            return;
        }

        var options = new PdfAnnotationRemovalOptions
        {
            ObjectNumber = ObjectNumber,
            PageNumber = PageNumber,
            Subtype = Subtype,
            RemoveMatchingPopups = !KeepPopups.IsPresent
        };

        PdfCommandUtilities.EnsureDirectory(outputPath);
        PdfAnnotationEditResult result = PdfAnnotationEditor.RemoveAnnotations(inputPath, outputPath, options);
        WriteObject(PassThruReport.IsPresent ? result : new FileInfo(outputPath));
    }
}
