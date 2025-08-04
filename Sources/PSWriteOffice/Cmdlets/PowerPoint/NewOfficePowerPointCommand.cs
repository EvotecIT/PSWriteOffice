using System;
using System.Management.Automation;
using ShapeCrawler;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

[Cmdlet(VerbsCommon.New, "OfficePowerPoint")]
public class NewOfficePowerPointCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public string FilePath { get; set; } = string.Empty;

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
