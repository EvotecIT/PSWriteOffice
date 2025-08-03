using System;
using System.IO;
using System.Management.Automation;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

[Cmdlet(VerbsCommon.Get, "OfficeWord")]
public class GetOfficeWordCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public string FilePath { get; set; } = string.Empty;

    [Parameter]
    public SwitchParameter ReadOnly { get; set; }

    [Parameter]
    public SwitchParameter AutoSave { get; set; }

    protected override void ProcessRecord()
    {
        try
        {
            var document = WordDocumentService.LoadDocument(FilePath, ReadOnly.IsPresent, AutoSave.IsPresent);
            WriteObject(document);
        }
        catch (FileNotFoundException ex)
        {
            WriteError(new ErrorRecord(ex, "FileNotFound", ErrorCategory.ObjectNotFound, FilePath));
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "WordLoadFailed", ErrorCategory.InvalidOperation, FilePath));
        }
    }
}
