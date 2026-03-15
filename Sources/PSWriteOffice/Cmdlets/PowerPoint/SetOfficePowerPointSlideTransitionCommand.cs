using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets the transition used when advancing to a slide.</summary>
/// <para>Works on an explicit slide object or on the current slide inside the PowerPoint DSL.</para>
/// <example>
///   <summary>Apply a fade transition to a slide.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Set-OfficePowerPointSlideTransition -Transition Fade</code>
///   <para>Updates the first slide so it uses the Fade transition.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointSlideTransition")]
[Alias("PptTransition")]
[OutputType(typeof(PowerPointSlide))]
public sealed class SetOfficePowerPointSlideTransitionCommand : PSCmdlet
{
    /// <summary>Slide to update (optional inside a slide DSL scope).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Transition to apply.</summary>
    [Parameter(Mandatory = true)]
    public SlideTransition Transition { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            var slide = Slide ?? PowerPointDslContext.Require(this).RequireSlide();
            slide.Transition = Transition;
            WriteObject(slide);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointSetTransitionFailed", ErrorCategory.InvalidOperation, Slide));
        }
    }
}
