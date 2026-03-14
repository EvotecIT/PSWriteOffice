using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds a Markdown horizontal rule.</summary>
/// <example>
///   <summary>Add a horizontal rule.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>MarkdownHorizontalRule</code>
///   <para>Appends a horizontal rule to the document.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownHorizontalRule", DefaultParameterSetName = ParameterSetContext)]
[Alias("MarkdownHorizontalRule", "MarkdownHr")]
[OutputType(typeof(MarkdownDoc))]
public sealed class AddOfficeMarkdownHorizontalRuleCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Emit the Markdown document after appending the rule.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var doc = ResolveDocument();
        doc.Hr();

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
