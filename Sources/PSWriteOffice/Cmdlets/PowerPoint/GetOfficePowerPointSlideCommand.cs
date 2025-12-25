using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Enumerates slides or retrieves a specific slide.</summary>
/// <para>Supports pipeline-friendly iteration over <see cref="PowerPointPresentation.Slides"/> or direct index selection.</para>
/// <example>
///   <summary>List slide titles.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointSlide -Presentation $ppt | ForEach-Object { $_.GetPlaceholder([DocumentFormat.OpenXml.Presentation.PlaceholderValues]::Title).Text }</code>
///   <para>Streams each slide so you can read the title placeholder text.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointSlide")]
public class GetOfficePowerPointSlideCommand : PSCmdlet
{
    /// <summary>Presentation to inspect.</summary>
    [Parameter(Mandatory = true)]
    public PowerPointPresentation Presentation { get; set; } = null!;

    /// <summary>Optional zero-based index; omit to enumerate all slides.</summary>
    [Parameter]
    public int? Index { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Index.HasValue)
        {
            if (Index.Value < 0 || Index.Value >= Presentation.Slides.Count)
            {
                WriteError(new ErrorRecord(new ArgumentOutOfRangeException(nameof(Index)), "PowerPointSlideIndexOutOfRange", ErrorCategory.InvalidArgument, Index));
                return;
            }

            WriteObject(Presentation.Slides[Index.Value]);
        }
        else
        {
            foreach (var slide in Presentation.Slides)
            {
                WriteObject(slide);
            }
        }
    }
}
