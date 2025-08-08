/*
.SYNOPSIS
Saves a Word document.

.DESCRIPTION
Saves the Word document to its current or a new file path and can optionally display the file.

.EXAMPLE
PS> .\Examples\Word\Save-OfficeWord.ps1
Shows how to save a document to disk.
*/
using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

[Cmdlet(VerbsData.Save, "OfficeWord")]
public class SaveOfficeWordCommand : PSCmdlet
{
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public WordDocument Document { get; set; } = null!;

    [Parameter]
    public SwitchParameter Show { get; set; }

    [Parameter]
    public string? FilePath { get; set; }

    protected override void ProcessRecord()
    {
        if (Document == null)
        {
            WriteError(new ErrorRecord(new ArgumentNullException(nameof(Document)), "DocumentNull", ErrorCategory.InvalidArgument, null));
            return;
        }

        try
        {
            WordDocumentService.SaveDocument(Document, Show.IsPresent, FilePath);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "WordSaveFailed", ErrorCategory.InvalidOperation, FilePath));
        }
    }
}
