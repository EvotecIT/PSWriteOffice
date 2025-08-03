using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

[Cmdlet(VerbsCommon.Close, "OfficeWord")]
public class CloseOfficeWordCommand : PSCmdlet
{
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public WordDocument Document { get; set; } = null!;

    protected override void ProcessRecord()
    {
        if (Document == null)
        {
            WriteError(new ErrorRecord(new ArgumentNullException(nameof(Document)), "DocumentNull", ErrorCategory.InvalidArgument, null));
            return;
        }

        try
        {
            WordDocumentService.CloseDocument(Document);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "WordCloseFailed", ErrorCategory.InvalidOperation, Document));
        }
    }
}
