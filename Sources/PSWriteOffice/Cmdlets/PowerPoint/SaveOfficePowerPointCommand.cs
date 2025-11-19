using System;
using System.Management.Automation;
using ShapeCrawler;
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
    public Presentation Presentation { get; set; } = null!;

    /// <summary>Launch the saved file in the default viewer.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

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
                PowerPointDocumentService.SavePresentation(Presentation, Show.IsPresent);
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointSaveFailed", ErrorCategory.InvalidOperation, null));
        }
    }
}
