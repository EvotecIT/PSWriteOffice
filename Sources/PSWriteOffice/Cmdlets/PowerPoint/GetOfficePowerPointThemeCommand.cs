using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Gets theme information for a PowerPoint presentation master.</summary>
/// <example>
///   <summary>Inspect the default master theme.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointTheme -Presentation $ppt</code>
///   <para>Returns the theme name, theme colors, and configured fonts for master 0.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointTheme")]
[Alias("PptTheme")]
[OutputType(typeof(PowerPointThemeInfo))]
public sealed class GetOfficePowerPointThemeCommand : PSCmdlet
{
    /// <summary>Presentation to inspect (optional inside New-OfficePowerPoint).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Slide master index to inspect.</summary>
    [Parameter]
    public int Master { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            var presentation = Presentation ?? PowerPointDslContext.Require(this).Presentation;
            var info = new PowerPointThemeInfo(
                Master,
                presentation.ThemeName,
                presentation.GetThemeColors(Master),
                presentation.GetThemeFonts(Master));

            WriteObject(info);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointGetThemeFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }
}
