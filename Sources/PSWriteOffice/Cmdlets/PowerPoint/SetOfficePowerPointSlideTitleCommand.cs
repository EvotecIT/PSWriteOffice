using System;
using System.Management.Automation;
using ShapeCrawler;

namespace PSWriteOffice.Cmdlets.PowerPoint;

[Cmdlet(VerbsCommon.Set, "OfficePowerPointSlideTitle")]
public class SetOfficePowerPointSlideTitleCommand : PSCmdlet
{
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ISlide Slide { get; set; } = null!;

    [Parameter(Mandatory = true)]
    public string Title { get; set; } = null!;

    protected override void ProcessRecord()
    {
        try
        {
            var titleShape = Slide.Shapes.Shape("Title 1");
            titleShape.TextBox!.SetText(Title);
            WriteObject(Slide);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointSetTitleFailed", ErrorCategory.InvalidOperation, Slide));
        }
    }
}
