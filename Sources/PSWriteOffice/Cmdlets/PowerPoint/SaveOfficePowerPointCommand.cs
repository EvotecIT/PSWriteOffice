using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Saves a presentation without disposing it.</summary>
/// <para>Use <c>Close-OfficePowerPoint -Save</c> when the presentation should be saved and closed.</para>
/// <example>
///   <summary>Save and open the deck.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ppt = New-OfficePowerPoint -FilePath .\Examples\Documents\PowerPointSave.pptx
/// $slide = Add-OfficePowerPointSlide -Presentation $ppt -Layout 1
/// Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Saved later'
/// Save-OfficePowerPoint -Presentation $ppt -PdfPath .\Examples\Documents\PowerPointSave.pdf</code>
///   <para>Saves the current presentation and exports a PDF sidecar.</para>
/// </example>
[Cmdlet(VerbsData.Save, "OfficePowerPoint", SupportsShouldProcess = true)]
[OutputType(typeof(PowerPointPresentation))]
public class SaveOfficePowerPointCommand : PSCmdlet
{
    /// <summary>Presentation instance to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    [ValidateNotNull]
    public PowerPointPresentation Presentation { get; set; } = null!;

    /// <summary>Optional save-as path.</summary>
    [Parameter]
    [Alias("FilePath")]
    public string? Path { get; set; }

    /// <summary>Launch the saved file in the default viewer.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <summary>Password used to save the presentation as an encrypted package.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>Optional PDF path to create from the same presentation.</summary>
    [Parameter]
    public string? PdfPath { get; set; }

    /// <summary>Emit the still-open presentation for further processing.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Presentation == null)
        {
            WriteError(new ErrorRecord(new ArgumentNullException(nameof(Presentation)), "PresentationNull", ErrorCategory.InvalidArgument, null));
            return;
        }

        try
        {
            var associatedPath = PowerPointDocumentService.GetAssociatedPath(Presentation);
            if (string.IsNullOrWhiteSpace(Path) && string.IsNullOrWhiteSpace(associatedPath))
            {
                throw new PSInvalidOperationException("No file path provided. Use -Path or open the presentation from disk.");
            }

            var targetPath = string.IsNullOrWhiteSpace(Path)
                ? associatedPath!
                : SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            if (ShouldProcess(targetPath, "Save PowerPoint presentation"))
            {
                PowerPointDocumentService.SavePresentation(Presentation, Show.IsPresent, Password, targetPath);
                SavePdfIfRequested();
                if (PassThru.IsPresent)
                {
                    WriteObject(Presentation);
                }
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointSaveFailed", ErrorCategory.InvalidOperation, null));
        }
    }

    private void SavePdfIfRequested()
    {
        if (string.IsNullOrWhiteSpace(PdfPath))
        {
            return;
        }

        var pdfPath = PdfCommandUtilities.ResolvePath(this, PdfPath!);
        if (!PdfCommandUtilities.ShouldWrite(this, pdfPath, "Write PowerPoint PDF"))
        {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(pdfPath);
        Presentation.SaveAsPdf(pdfPath).RequireSuccess();
    }
}
