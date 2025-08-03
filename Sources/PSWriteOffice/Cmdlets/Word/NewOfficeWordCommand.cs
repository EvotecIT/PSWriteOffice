using System;
using System.Management.Automation;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

[Cmdlet(VerbsCommon.New, "OfficeWord")]
public class NewOfficeWordCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public string FilePath { get; set; } = string.Empty;

    [Parameter]
    public SwitchParameter AutoSave { get; set; }

    protected override void ProcessRecord()
    {
        try
        {
            var document = WordDocumentService.CreateDocument(FilePath, AutoSave.IsPresent);
            WriteObject(document);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "WordCreateFailed", ErrorCategory.InvalidOperation, FilePath));
        }
    }
}
