using System;
using System.Management.Automation;
using ShapeCrawler;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

[Cmdlet(VerbsData.Save, "OfficePowerPoint", SupportsShouldProcess = true)]
public class SaveOfficePowerPointCommand : PSCmdlet
{
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    [ValidateNotNull]
    public Presentation Presentation { get; set; } = null!;

    [Parameter]
    public SwitchParameter Show { get; set; }

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
