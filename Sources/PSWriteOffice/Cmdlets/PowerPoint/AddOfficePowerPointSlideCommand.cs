using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a new slide to a PowerPoint presentation.</summary>
/// <para>Creates a slide using OfficeIMO master and layout indexes.</para>
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
    public PowerPointPresentation Presentation { get; set; } = null!;

    /// <summary>Slide master index to use.</summary>
    [Parameter]
    public int Master { get; set; } = 0;

    /// <summary>Layout index to use (matches the template’s built-in layouts).</summary>
    [Parameter]
    public int Layout { get; set; } = 1;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            var slide = Presentation.AddSlide(Master, Layout);
            WriteObject(slide);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointAddSlideFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }
}
