using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds a Markdown block quote.</summary>
/// <example>
///   <summary>Add a quote block.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>MarkdownQuote -Text 'Key takeaway goes here.'</code>
///   <para>Appends a quote block to the document.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownQuote", DefaultParameterSetName = ParameterSetContext)]
[Alias("MarkdownQuote")]
[OutputType(typeof(MarkdownDoc))]
public sealed class AddOfficeMarkdownQuoteCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Quote text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Emit the Markdown document after appending the quote.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var doc = ResolveDocument();
        doc.Quote(Text ?? string.Empty);

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
