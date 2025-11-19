using System;
using System.Management.Automation;
using ShapeCrawler;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a new slide to a PowerPoint presentation.</summary>
/// <para>Wraps ShapeCrawler to append a slide using a built-in layout index.</para>
/// <example>
///   <summary>Append a slide with the default layout.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ppt = New-OfficePowerPoint -FilePath .\deck.pptx; Add-OfficePowerPointSlide -Presentation $ppt</code>
///   <para>Creates a deck and appends a new slide at the end.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointSlide")]
public class AddOfficePowerPointSlideCommand : PSCmdlet
{
    /// <summary>Presentation to update.</summary>
    [Parameter(Mandatory = true)]
    public Presentation Presentation { get; set; } = null!;

    /// <summary>Layout index to use (matches the template’s built-in layouts).</summary>
    [Parameter]
    public int Layout { get; set; } = 1;

    /// <inheritdoc />
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
