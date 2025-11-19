using System;
using System.Management.Automation;
using ShapeCrawler;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Merges slides from other decks into the target presentation.</summary>
/// <para>Uses ShapeCrawler to append each slide sequentially.</para>
/// <example>
///   <summary>Combine multiple decks.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Merge-OfficePowerPoint -Presentation $ppt -FilePath '.\Intro.pptx','.\Appendix.pptx'</code>
///   <para>Appends all slides from Intro and Appendix into <c>$ppt</c>.</para>
/// </example>
[Cmdlet(VerbsData.Merge, "OfficePowerPoint")]
public class MergeOfficePowerPointCommand : PSCmdlet
{
    /// <summary>Destination presentation that receives the extra slides.</summary>
    [Parameter(Mandatory = true)]
    public Presentation Presentation { get; set; } = null!;

    /// <summary>Paths to source decks whose slides should be appended.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string[] FilePath { get; set; } = Array.Empty<string>();

    /// <inheritdoc />
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
