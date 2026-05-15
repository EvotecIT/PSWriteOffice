using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Closes a PowerPoint presentation and optionally saves it.</summary>
/// <para>Provides a cmdlet wrapper so PowerShell scripts do not need to call <c>Dispose</c> directly.</para>
/// <example>
///   <summary>Close without saving.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ppt = Get-OfficePowerPoint -FilePath .\deck.pptx; Close-OfficePowerPoint -Presentation $ppt</code>
///   <para>Releases the loaded presentation instance.</para>
/// </example>
/// <example>
///   <summary>Save, open, and close.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Close-OfficePowerPoint -Presentation $ppt -Save -Show</code>
///   <para>Saves the presentation, opens it in PowerPoint, and releases the object.</para>
/// </example>
[Cmdlet(VerbsCommon.Close, "OfficePowerPoint", SupportsShouldProcess = true)]
public sealed class CloseOfficePowerPointCommand : PSCmdlet
{
    /// <summary>Presentation to close.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    [ValidateNotNull]
    public PowerPointPresentation Presentation { get; set; } = null!;

    /// <summary>Persist changes before closing.</summary>
    [Parameter]
    public SwitchParameter Save { get; set; }

    /// <summary>Open the presentation in PowerPoint after saving.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Presentation == null)
        {
            WriteError(new ErrorRecord(new ArgumentNullException(nameof(Presentation)), "PresentationNull", ErrorCategory.InvalidArgument, null));
            return;
        }

        try
        {
            var action = Save.IsPresent || Show.IsPresent ? "Save and close" : "Close";
            if (ShouldProcess("PowerPoint presentation", action))
            {
                PowerPointDocumentService.ClosePresentation(Presentation, Save.IsPresent, Show.IsPresent);
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointCloseFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }
}
