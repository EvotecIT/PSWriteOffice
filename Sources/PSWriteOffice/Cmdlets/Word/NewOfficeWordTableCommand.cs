/*
.SYNOPSIS
Adds a table to a Word document.

.DESCRIPTION
Creates a table in a Word document from an array of objects with optional style and layout.

.EXAMPLE
PS> .\Examples\Word\New-OfficeWordTable.ps1
Shows how to insert a table into a document.
*/
using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

[Cmdlet(VerbsCommon.New, "OfficeWordTable")]
public class NewOfficeWordTableCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public WordDocument Document { get; set; } = null!;

    [Parameter(Mandatory = true)]
    public Array DataTable { get; set; } = Array.Empty<object>();

    [Parameter]
    public WordTableStyle Style { get; set; } = WordTableStyle.TableGrid;

    [Parameter]
    [ValidateSet("Autofit", "Fixed")]
    public string? TableLayout { get; set; }

    [Parameter]
    public SwitchParameter SkipHeader { get; set; }

    [Parameter]
    public SwitchParameter Suppress { get; set; }

    protected override void ProcessRecord()
    {
        if (Document == null)
        {
            WriteError(new ErrorRecord(new ArgumentNullException(nameof(Document)), "DocumentNull", ErrorCategory.InvalidArgument, null));
            return;
        }

        var table = WordDocumentService.AddTable(Document, DataTable, Style, TableLayout, SkipHeader);

        if (!Suppress)
        {
            WriteObject(table);
        }
    }
}
