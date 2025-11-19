using System;
using System.Management.Automation;
using ShapeCrawler;
using PSWriteOffice;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Removes a slide by index.</summary>
/// <para>Supports <c>-WhatIf</c>/<c>-Confirm</c> thanks to <c>SupportsShouldProcess</c>.</para>
/// <example>
///   <summary>Delete the first slide.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Remove-OfficePowerPointSlide -Presentation $ppt -Index 0</code>
///   <para>Removes slide 1 from the deck.</para>
/// </example>
[Cmdlet(VerbsCommon.Remove, "OfficePowerPointSlide", SupportsShouldProcess = true)]
public class RemoveOfficePowerPointSlideCommand : PSCmdlet
{
    /// <summary>Presentation to modify.</summary>
    [Parameter(Mandatory = true)]
    public Presentation Presentation { get; set; } = null!;

    /// <summary>Zero-based slide index.</summary>
    [Parameter(Mandatory = true)]
    public int Index { get; set; }

    /// <inheritdoc />
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
