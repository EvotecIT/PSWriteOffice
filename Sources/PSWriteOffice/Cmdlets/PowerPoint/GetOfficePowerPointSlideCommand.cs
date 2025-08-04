using System;
using System.Management.Automation;
using ShapeCrawler;

namespace PSWriteOffice.Cmdlets.PowerPoint;

[Cmdlet(VerbsCommon.Get, "OfficePowerPointSlide")]
public class GetOfficePowerPointSlideCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public Presentation Presentation { get; set; } = null!;

    [Parameter]
    public int? Index { get; set; }

    protected override void ProcessRecord()
    {
        if (Index.HasValue)
        {
            if (Index.Value < 0 || Index.Value >= Presentation.Slides.Count)
            {
                WriteError(new ErrorRecord(new ArgumentOutOfRangeException(nameof(Index)), "PowerPointSlideIndexOutOfRange", ErrorCategory.InvalidArgument, Index));
                return;
            }

            WriteObject(Presentation.Slides[Index.Value]);
        }
        else
        {
            foreach (var slide in Presentation.Slides)
            {
                WriteObject(slide);
            }
        }
    }
}
