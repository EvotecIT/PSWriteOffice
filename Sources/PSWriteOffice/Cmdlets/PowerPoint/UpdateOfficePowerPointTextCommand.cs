using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Replaces text in a PowerPoint slide or presentation.</summary>
/// <para>Can replace text in text boxes, tables, and optionally notes using the OfficeIMO text replacement helpers.</para>
/// <example>
///   <summary>Replace fiscal year text across the whole deck.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Update-OfficePowerPointText -Presentation $ppt -OldValue 'FY24' -NewValue 'FY25' -IncludeNotes</code>
///   <para>Replaces matching text throughout the presentation and notes.</para>
/// </example>
[Cmdlet(VerbsData.Update, "OfficePowerPointText", DefaultParameterSetName = ParameterSetAuto)]
[Alias("Replace-OfficePowerPointText")]
[OutputType(typeof(int))]
public sealed class UpdateOfficePowerPointTextCommand : PSCmdlet
{
    private const string ParameterSetAuto = "Auto";
    private const string ParameterSetPresentation = "Presentation";
    private const string ParameterSetSlide = "Slide";

    /// <summary>Presentation to update.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetPresentation)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Slide to update.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetSlide)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Text to find.</summary>
    [Parameter(Mandatory = true)]
    public string OldValue { get; set; } = string.Empty;

    /// <summary>Replacement text.</summary>
    [Parameter(Mandatory = true)]
    [AllowNull]
    public string? NewValue { get; set; }

    /// <summary>Include table cells in the replacement operation.</summary>
    [Parameter]
    public bool IncludeTables { get; set; } = true;

    /// <summary>Include notes text in the replacement operation.</summary>
    [Parameter]
    public SwitchParameter IncludeNotes { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            int replacements;
            switch (ParameterSetName)
            {
                case ParameterSetPresentation:
                    replacements = (Presentation ?? throw new InvalidOperationException("Presentation was not provided."))
                        .ReplaceText(OldValue, NewValue ?? string.Empty, IncludeTables, IncludeNotes.IsPresent);
                    break;
                case ParameterSetSlide:
                    replacements = (Slide ?? throw new InvalidOperationException("Slide was not provided."))
                        .ReplaceText(OldValue, NewValue ?? string.Empty, IncludeTables, IncludeNotes.IsPresent);
                    break;
                default:
                    replacements = ReplaceUsingDslContext();
                    break;
            }

            WriteObject(replacements);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointUpdateTextFailed", ErrorCategory.InvalidOperation, Presentation ?? (object?)Slide));
        }
    }

    private int ReplaceUsingDslContext()
    {
        var context = PowerPointDslContext.Current
            ?? throw new InvalidOperationException("Specify -Presentation, -Slide, or run inside New-OfficePowerPoint.");

        if (context.CurrentSlide != null)
        {
            return context.CurrentSlide.ReplaceText(OldValue, NewValue ?? string.Empty, IncludeTables, IncludeNotes.IsPresent);
        }

        return context.Presentation.ReplaceText(OldValue, NewValue ?? string.Empty, IncludeTables, IncludeNotes.IsPresent);
    }
}
