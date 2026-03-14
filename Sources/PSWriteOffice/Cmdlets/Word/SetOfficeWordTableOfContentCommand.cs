using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Sets properties on a table of contents in a Word document.</summary>
/// <example>
///   <summary>Update the table of contents headings.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Set-OfficeWordTableOfContent -Text 'Contents' -TextNoContent 'No entries'</code>
///   <para>Updates the table of contents display text.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeWordTableOfContent")]
[OutputType(typeof(WordTableOfContent))]
public sealed class SetOfficeWordTableOfContentCommand : PSCmdlet
{
    private const string ParameterSetTableOfContent = "TableOfContent";
    private const string ParameterSetDocument = "Document";

    /// <summary>Table of contents to update.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetTableOfContent)]
    public WordTableOfContent? TableOfContent { get; set; }

    /// <summary>Document to update when provided explicitly.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument? Document { get; set; }

    /// <summary>Heading text for the table of contents.</summary>
    [Parameter]
    public string? Text { get; set; }

    /// <summary>Text shown when the table of contents has no entries.</summary>
    [Parameter]
    public string? TextNoContent { get; set; }

    /// <summary>Emit the updated table of contents.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Text == null && TextNoContent == null)
        {
            throw new PSArgumentException("Specify -Text or -TextNoContent to update the table of contents.");
        }

        var toc = ResolveTableOfContent();

        if (Text != null)
        {
            toc.Text = Text;
        }

        if (TextNoContent != null)
        {
            toc.TextNoContent = TextNoContent;
        }

        if (PassThru.IsPresent)
        {
            WriteObject(toc);
        }
    }

    private WordTableOfContent ResolveTableOfContent()
    {
        if (TableOfContent != null)
        {
            return TableOfContent;
        }

        var document = Document ?? WordDslContext.Require(this).Document;
        if (document == null)
        {
            throw new InvalidOperationException("Word document was not provided.");
        }

        var toc = document.TableOfContent;
        if (toc == null)
        {
            throw new InvalidOperationException("Table of contents was not found in the document.");
        }

        return toc;
    }
}
