using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Removes a slide by index.</summary>
/// <para>Supports <c>-WhatIf</c>/<c>-Confirm</c> thanks to <c>SupportsShouldProcess</c>.</para>
/// <example>
///   <summary>Delete the first slide.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ppt = New-OfficePowerPoint -FilePath .\Examples\Documents\PowerPointRemoveSlide.pptx
/// Add-OfficePowerPointSlide -Presentation $ppt -Layout 1 | Out-Null
/// Add-OfficePowerPointSlide -Presentation $ppt -Layout 1 | Out-Null
/// Remove-OfficePowerPointSlide -Presentation $ppt -Index 0 -Confirm:$false
/// Save-OfficePowerPoint -Presentation $ppt</code>
///   <para>Removes the first slide and saves the updated deck.</para>
/// </example>
[Cmdlet(VerbsCommon.Remove, "OfficePowerPointSlide", SupportsShouldProcess = true)]
public class RemoveOfficePowerPointSlideCommand : PSCmdlet
{
    /// <summary>Presentation to modify.</summary>
    [Parameter(Mandatory = true)]
    public PowerPointPresentation Presentation { get; set; } = null!;

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
                Presentation.RemoveSlide(Index);
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointRemoveSlideFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }
}
