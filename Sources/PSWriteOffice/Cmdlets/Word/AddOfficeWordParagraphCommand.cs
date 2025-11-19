using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a paragraph to the current section/header/footer context.</summary>
/// <para>Acts as the primary DSL container for inline content such as text runs, bold segments, and images.</para>
/// <example>
///   <summary>Write a formatted sentence.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph { Add-OfficeWordText -Text 'Hello '; Add-OfficeWordText -Text 'World' -Bold }</code>
///   <para>Outputs “Hello World” with the second word bolded.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordParagraph")]
[Alias("WordParagraph")]
public sealed class AddOfficeWordParagraphCommand : PSCmdlet
{
    [Parameter(Position = 0)]
    public string? Text { get; set; }

    [Parameter]
    public ScriptBlock? Content { get; set; }

    [Parameter]
    public JustificationValues? Alignment { get; set; }

    [Parameter]
    public WordParagraphStyles? Style { get; set; }

    [Parameter]
    public SwitchParameter PassThru { get; set; }

    protected override void ProcessRecord()
    {
        var context = WordDslContext.Require(this);
        var host = context.RequireParagraphHost();
        var paragraph = host.AddParagraph(Text);

        if (Alignment.HasValue)
        {
            paragraph.ParagraphAlignment = Alignment.Value;
        }

        if (Style.HasValue)
        {
            paragraph.Style = Style.Value;
        }

        using (context.Push(paragraph))
        {
            Content?.InvokeReturnAsIs();
        }

        if (PassThru.IsPresent)
        {
            WriteObject(paragraph);
        }
    }
}
