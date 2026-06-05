using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Saves a presentation to disk.</summary>
/// <para>Invokes the PowerPoint service to persist the document and optionally launch it.</para>
/// <example>
///   <summary>Save and open the deck.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Save-OfficePowerPoint -Presentation $ppt -Show</code>
///   <para>Saves the current presentation and opens it in PowerPoint.</para>
/// </example>
[Cmdlet(VerbsData.Save, "OfficePowerPoint", SupportsShouldProcess = true)]
public class SaveOfficePowerPointCommand : PSCmdlet
{
    /// <summary>Presentation instance to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    [ValidateNotNull]
    public PowerPointPresentation Presentation { get; set; } = null!;

    /// <summary>Launch the saved file in the default viewer.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <summary>Password used to save the presentation as an encrypted package.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>Optional PDF path to create from the same presentation.</summary>
    [Parameter]
    public string? PdfPath { get; set; }

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
            if (ShouldProcess("PowerPoint presentation", "Save"))
            {
                SavePdfIfRequested();
                PowerPointDocumentService.SavePresentation(Presentation, Show.IsPresent, Password);
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

        Presentation.SaveAsPdf(PdfCommandUtilities.ResolvePath(this, PdfPath!));
    }
}
