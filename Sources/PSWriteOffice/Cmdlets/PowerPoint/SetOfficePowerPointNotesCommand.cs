using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets speaker notes for a PowerPoint slide.</summary>
/// <example>
///   <summary>Attach speaker notes.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Set-OfficePowerPointNotes -Slide $slide -Text 'Keep this under five minutes.'</code>
///   <para>Writes notes to the slide.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointNotes")]
[Alias("PptNotes")]
[OutputType(typeof(PowerPointSlide))]
public sealed class SetOfficePowerPointNotesCommand : PSCmdlet
{
    /// <summary>Slide whose notes should be updated (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Notes text to apply.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            var slide = Slide ?? PowerPointDslContext.Require(this).RequireSlide();
            var notes = slide.Notes;
            notes.Text = Text ?? string.Empty;
            WriteObject(slide);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointSetNotesFailed", ErrorCategory.InvalidOperation, Slide));
        }
    }
}
