using System;
using System.Management.Automation;
using ShapeCrawler;
using PSWriteOffice;

namespace PSWriteOffice.Cmdlets.PowerPoint;

[Cmdlet(VerbsCommon.Remove, "OfficePowerPointSlide", SupportsShouldProcess = true)]
public class RemoveOfficePowerPointSlideCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public Presentation Presentation { get; set; } = null!;

    [Parameter(Mandatory = true)]
    public int Index { get; set; }

    protected override void ProcessRecord()
    {
        try
        {
            if (ShouldProcess($"Slide {Index}", "Remove slide"))
            {
                Presentation.Slides.RemoveAt(Index);
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointRemoveSlideFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }
}
