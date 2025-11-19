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
    /// <summary>Optional initial paragraph text.</summary>
    [Parameter(Position = 0)]
    public string? Text { get; set; }

    /// <summary>Nested DSL content (runs, lists, images).</summary>
    [Parameter]
    public ScriptBlock? Content { get; set; }

    /// <summary>Paragraph justification.</summary>
    [Parameter]
    public JustificationValues? Alignment { get; set; }

    /// <summary>Paragraph style.</summary>
    [Parameter]
    public WordParagraphStyles? Style { get; set; }

    /// <summary>Emit the <see cref="WordParagraph"/> for further use.</summary>
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
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
