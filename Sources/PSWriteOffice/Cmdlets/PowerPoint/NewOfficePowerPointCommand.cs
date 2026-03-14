using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Creates a PowerPoint presentation using the DSL.</summary>
/// <para>Initializes a presentation, runs the DSL script block, and optionally saves the deck.</para>
/// <example>
///   <summary>Create and capture the presentation object.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ppt = New-OfficePowerPoint -FilePath .\deck.pptx</code>
///   <para>Creates <c>deck.pptx</c> and returns the live presentation object for further editing.</para>
/// </example>
/// <example>
///   <summary>Create a deck with a title slide.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\deck.pptx { PptSlide { PptTitle -Title 'Status Update' } } -Open</code>
///   <para>Creates, saves, and opens a deck with one titled slide.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePowerPoint")]
public class NewOfficePowerPointCommand : PSCmdlet
{
    /// <summary>Destination path for the new .pptx.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("Path")]
    public string FilePath { get; set; } = string.Empty;

    /// <summary>DSL scriptblock describing presentation content.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Open the presentation after saving.</summary>
    [Parameter]
    public SwitchParameter Open { get; set; }

    /// <summary>Skip saving after executing the DSL.</summary>
    [Parameter]
    public SwitchParameter NoSave { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> for chaining.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(FilePath);
        var directory = Path.GetDirectoryName(resolvedPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        PowerPointPresentation? presentation = null;
        try
        {
            presentation = PowerPointDocumentService.CreatePresentation(resolvedPath);

            if (Content == null)
            {
                WriteObject(presentation);
                return;
            }

            using (PowerPointDslContext.Enter(presentation))
            {
                try
                {
                    Content.InvokeReturnAsIs();
                }
                catch (Exception ex) when (IsStopUpstream(ex))
                {
                    // Select-Object -First throws this to stop enumeration; ignore.
                }
            }

            if (NoSave.IsPresent)
            {
                presentation.Dispose();
                return;
            }

            PowerPointDocumentService.SavePresentation(presentation, Open.IsPresent);

            if (PassThru.IsPresent)
            {
                WriteObject(new FileInfo(resolvedPath));
            }
        }
        catch (Exception ex)
        {
            presentation?.Dispose();
            WriteError(new ErrorRecord(ex, "PowerPointCreateFailed", ErrorCategory.InvalidOperation, FilePath));
        }
    }

    private static bool IsStopUpstream(Exception ex)
    {
        return ex.GetType().Name == "StopUpstreamCommandsException";
    }
}
