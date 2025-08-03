using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

[Cmdlet(VerbsCommon.New, "OfficeWordList")]
public class NewOfficeWordListCommand : PSCmdlet
{
    [Parameter]
    public ScriptBlock? Content { get; set; }

    [Parameter(Mandatory = true)]
    public WordDocument Document { get; set; } = null!;

    [Parameter]
    public WordListStyle Style { get; set; } = WordListStyle.Bulleted;

    [Parameter]
    public SwitchParameter Suppress { get; set; }

    protected override void ProcessRecord()
    {
        var list = WordDocumentService.AddList(Document, Style);

        if (Content != null)
        {
            var results = Content.Invoke();
            foreach (var result in results)
            {
                if (result.BaseObject is System.Collections.IDictionary dict)
                {
                    var level = dict["Level"] is int l ? l : 0;
                    var text = dict["Text"] as string[] ?? Array.Empty<string>();
                    WordDocumentService.AddListItem(list, level, text);
                }
            }
        }

        if (!Suppress)
        {
            WriteObject(list);
        }
    }
}
