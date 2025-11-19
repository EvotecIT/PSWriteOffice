using System;
using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds inline text to the current paragraph.</summary>
/// <para>Supports bold/italic/underline and color tweaks for quick DSL composition.</para>
/// <example>
///   <summary>Append bold text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph { Add-OfficeWordText -Text 'Important: ' -Bold }</code>
///   <para>Writes “Important:” with bold formatting.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordText")]
[Alias("WordText", "WordBold", "WordItalic")]
public sealed class AddOfficeWordTextCommand : PSCmdlet
{
    [Parameter(Mandatory = true, Position = 0)]
    public string[] Text { get; set; } = Array.Empty<string>();

    [Parameter]
    public SwitchParameter Bold { get; set; }

    [Parameter]
    public SwitchParameter Italic { get; set; }

    [Parameter]
    public UnderlineValues? Underline { get; set; }

    [Parameter]
    public string? Color { get; set; }

    protected override void BeginProcessing()
    {
        var name = MyInvocation.InvocationName;
        if (string.Equals(name, "WordBold", StringComparison.OrdinalIgnoreCase))
        {
            Bold = true;
        }
        else if (string.Equals(name, "WordItalic", StringComparison.OrdinalIgnoreCase))
        {
            Italic = true;
        }
    }

    protected override void ProcessRecord()
    {
        var context = WordDslContext.Require(this);
        var paragraph = context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();

        foreach (var entry in Text)
        {
            var run = paragraph.AddFormattedText(entry, Bold.IsPresent, Italic.IsPresent, Underline);
            if (!string.IsNullOrWhiteSpace(Color))
            {
                run.SetColorHex(Color);
            }
        }
    }
}
