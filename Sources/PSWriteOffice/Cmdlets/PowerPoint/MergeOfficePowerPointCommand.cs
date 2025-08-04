using System;
using System.Management.Automation;
using ShapeCrawler;

namespace PSWriteOffice.Cmdlets.PowerPoint;

[Cmdlet(VerbsData.Merge, "OfficePowerPoint")]
public class MergeOfficePowerPointCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public Presentation Presentation { get; set; } = null!;

    [Parameter(Mandatory = true, Position = 1)]
    public string[] FilePath { get; set; } = Array.Empty<string>();

    protected override void ProcessRecord()
    {
        try
        {
            foreach (var path in FilePath)
            {
                using var source = new Presentation(path);
                foreach (var slide in source.Slides)
                {
                    Presentation.Slides.Add(slide);
                }
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointMergeFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }
}
