using System;
using System.IO;
using System.Management.Automation;
using ShapeCrawler;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Loads an existing PowerPoint presentation.</summary>
/// <para>Returns a ShapeCrawler <see cref="Presentation"/> for downstream slide operations.</para>
/// <example>
///   <summary>Open a deck for editing.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ppt = Get-OfficePowerPoint -FilePath .\Quarterly.pptx</code>
///   <para>Reads <c>Quarterly.pptx</c> and exposes the presentation object.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPoint")]
public class GetOfficePowerPointCommand : PSCmdlet
{
    /// <summary>Path to the .pptx file.</summary>
    [Parameter(Mandatory = true)]
    [ValidateNotNullOrEmpty]
    public string FilePath { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            var presentation = PowerPointDocumentService.LoadPresentation(FilePath);
            WriteObject(presentation);
        }
        catch (FileNotFoundException ex)
        {
            WriteError(new ErrorRecord(ex, "FileNotFound", ErrorCategory.ObjectNotFound, FilePath));
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointLoadFailed", ErrorCategory.InvalidOperation, FilePath));
        }
    }
}
