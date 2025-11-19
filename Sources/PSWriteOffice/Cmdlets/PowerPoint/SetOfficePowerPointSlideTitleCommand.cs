using System;
using System.Management.Automation;
using ShapeCrawler;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets the text of the title placeholder on a slide.</summary>
/// <para>Targets the default “Title 1” shape created by most master layouts.</para>
/// <example>
///   <summary>Rename a slide.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Set-OfficePowerPointSlideTitle -Title 'Executive Summary'</code>
///   <para>Updates the first slide’s title.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointSlideTitle")]
public class SetOfficePowerPointSlideTitleCommand : PSCmdlet
{
    /// <summary>Slide whose title should change.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ISlide Slide { get; set; } = null!;

    /// <summary>New title text.</summary>
    [Parameter(Mandatory = true)]
    public string Title { get; set; } = null!;

    /// <inheritdoc />
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
