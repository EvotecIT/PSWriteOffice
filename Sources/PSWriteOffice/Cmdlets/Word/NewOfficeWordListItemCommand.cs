using System;
using System.Collections.Specialized;
using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

[Cmdlet(VerbsCommon.New, "OfficeWordListItem")]
public class NewOfficeWordListItemCommand : PSCmdlet
{
    [Parameter]
    public WordList? List { get; set; }

    [Parameter]
    public int Level { get; set; }

    [Parameter]
    public string[] Text { get; set; } = Array.Empty<string>();

    [Parameter]
    public bool?[]? Bold { get; set; }

    [Parameter]
    public bool?[]? Italic { get; set; }

    [Parameter]
    public UnderlineValues?[]? Underline { get; set; }

    [Parameter]
    public string[]? Color { get; set; }

    [Parameter]
    public JustificationValues? Alignment { get; set; }

    [Parameter]
    public SwitchParameter Suppress { get; set; }

    protected override void ProcessRecord()
    {
        if (List != null)
        {
            var item = WordDocumentService.AddListItem(List, Level, Text);
            if (!Suppress)
            {
                WriteObject(item);
            }
        }
        else
        {
            var ordered = new OrderedDictionary
            {
                { "List", null },
                { "Level", Level },
                { "Text", Text },
                { "Bold", Bold },
                { "Italic", Italic },
                { "Underline", Underline },
                { "Color", Color },
                { "Alignment", Alignment },
                { "Suppress", Suppress.IsPresent }
            };
            WriteObject(ordered);
        }
    }
}
