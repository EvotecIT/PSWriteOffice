using System;
using System.Management.Automation;
using ShapeCrawler;

namespace PSWriteOffice.Cmdlets.PowerPoint;

[Cmdlet(VerbsCommon.Add, "OfficePowerPointSlide")]
public class AddOfficePowerPointSlideCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public Presentation Presentation { get; set; } = null!;

    [Parameter]
    public int Layout { get; set; } = 1;

    protected override void ProcessRecord()
    {
        try
        {
            Presentation.Slides.Add(Layout);
            var slide = Presentation.Slides[Presentation.Slides.Count - 1];
            WriteObject(slide);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointAddSlideFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }
}
