using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Copies an existing slide within a PowerPoint presentation.</summary>
/// <para>Uses OfficeIMO slide duplication so charts, notes, and shapes are preserved.</para>
/// <example>
///   <summary>Duplicate the first slide and insert the copy after it.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Copy-OfficePowerPointSlide -Presentation $ppt -Index 0</code>
///   <para>Creates a duplicate of slide 1 and inserts it at position 2.</para>
/// </example>
[Cmdlet(VerbsCommon.Copy, "OfficePowerPointSlide")]
[OutputType(typeof(PowerPointSlide))]
public sealed class CopyOfficePowerPointSlideCommand : PSCmdlet
{
    /// <summary>Presentation to update (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Zero-based slide index to duplicate.</summary>
    [Parameter(Mandatory = true)]
    public int Index { get; set; }

    /// <summary>Optional target index for the duplicate; omit to insert after the source slide.</summary>
    [Parameter]
    public int? InsertAt { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            var presentation = Presentation ?? PowerPointDslContext.Current?.Presentation
                ?? throw new InvalidOperationException("Presentation was not provided. Use -Presentation or run inside New-OfficePowerPoint.");

            WriteObject(presentation.DuplicateSlide(Index, InsertAt));
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointCopySlideFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }
}
