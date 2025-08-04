using System;
using System.Management.Automation;
using ShapeCrawler;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

[Cmdlet(VerbsData.Save, "OfficePowerPoint")]
public class SaveOfficePowerPointCommand : PSCmdlet
{
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
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
            PowerPointDocumentService.SavePresentation(Presentation, Show.IsPresent);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointSaveFailed", ErrorCategory.InvalidOperation, null));
        }
    }
}
