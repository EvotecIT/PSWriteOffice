using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Creates a blank PowerPoint presentation.</summary>
/// <para>Initializes an OfficeIMO <see cref="PowerPointPresentation"/> backed by the supplied file path.</para>
/// <example>
///   <summary>Create and capture the presentation object.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ppt = New-OfficePowerPoint -FilePath .\deck.pptx</code>
///   <para>Creates <c>deck.pptx</c> and returns the live presentation object for further editing.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePowerPoint")]
public class NewOfficePowerPointCommand : PSCmdlet
{
    /// <summary>Destination path for the new .pptx.</summary>
    [Parameter(Mandatory = true)]
    public string FilePath { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            var presentation = PowerPointDocumentService.CreatePresentation(FilePath);
            WriteObject(presentation);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointCreateFailed", ErrorCategory.InvalidOperation, FilePath));
        }
    }
}
