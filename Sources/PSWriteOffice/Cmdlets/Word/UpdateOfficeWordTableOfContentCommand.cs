using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Updates the table of contents in a Word document.</summary>
/// <example>
///   <summary>Mark table of contents for refresh on open.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Update-OfficeWordTableOfContent</code>
///   <para>Marks TOC fields as dirty and updates the document settings.</para>
/// </example>
[Cmdlet(VerbsData.Update, "OfficeWordTableOfContent")]
[OutputType(typeof(WordTableOfContent))]
public sealed class UpdateOfficeWordTableOfContentCommand : PSCmdlet
{
    private const string ParameterSetTableOfContent = "TableOfContent";
    private const string ParameterSetDocument = "Document";

    /// <summary>Table of contents to update.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetTableOfContent)]
    public WordTableOfContent? TableOfContent { get; set; }

    /// <summary>Document to update when provided explicitly.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument? Document { get; set; }

    /// <summary>Rebuild the table of contents before updating.</summary>
    [Parameter]
    public SwitchParameter Regenerate { get; set; }

    /// <summary>Emit the updated table of contents.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = Document ?? WordDslContext.Require(this).Document;
        var toc = TableOfContent;

        if (toc == null && document == null)
        {
            throw new InvalidOperationException("Word document was not provided.");
        }

        if (Regenerate.IsPresent)
        {
            toc = toc != null ? toc.Regenerate() : document!.RegenerateTableOfContent();
        }
        else
        {
            if (toc == null)
            {
                toc = document!.TableOfContent;
            }

            if (toc == null)
            {
                throw new InvalidOperationException("Table of contents was not found in the document.");
            }

            toc.Update();
        }

        if (PassThru.IsPresent && toc != null)
        {
            WriteObject(toc);
        }
    }
}
