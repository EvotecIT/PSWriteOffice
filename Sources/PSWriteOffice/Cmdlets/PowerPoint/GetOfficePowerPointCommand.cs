using System;
using System.IO;
using System.Management.Automation;
using ShapeCrawler;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

[Cmdlet(VerbsCommon.Get, "OfficePowerPoint")]
public class GetOfficePowerPointCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    [ValidateNotNullOrEmpty]
    public string FilePath { get; set; } = string.Empty;

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
