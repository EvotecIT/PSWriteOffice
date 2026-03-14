using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds a Markdown paragraph.</summary>
/// <example>
///   <summary>Add a paragraph.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>MarkdownParagraph -Text 'This report is generated automatically.'</code>
///   <para>Appends a paragraph to the current Markdown document.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownParagraph", DefaultParameterSetName = ParameterSetContext)]
[Alias("MarkdownParagraph")]
[OutputType(typeof(MarkdownDoc))]
public sealed class AddOfficeMarkdownParagraphCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Paragraph text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Emit the Markdown document after appending the paragraph.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var doc = ResolveDocument();
        doc.P(Text ?? string.Empty);

        if (PassThru.IsPresent)
        {
            WriteObject(doc);
        }
    }

    private MarkdownDoc ResolveDocument()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            return Document ?? throw new PSArgumentException("Provide a Markdown document.");
        }

        return MarkdownDslContext.Require(this).Document;
    }
}
