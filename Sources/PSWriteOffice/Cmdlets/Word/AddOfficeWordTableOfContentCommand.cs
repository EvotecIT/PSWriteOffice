using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a table of contents to a Word document.</summary>
/// <example>
///   <summary>Add a default table of contents.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordTableOfContent</code>
///   <para>Inserts a table of contents using the default template.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordTableOfContent")]
[Alias("WordTableOfContent")]
[OutputType(typeof(WordTableOfContent))]
public sealed class AddOfficeWordTableOfContentCommand : PSCmdlet
{
    /// <summary>Document to modify when provided explicitly.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordDocument? Document { get; set; }

    /// <summary>Table of contents template style.</summary>
    [Parameter]
    public TableOfContentStyle Style { get; set; } = TableOfContentStyle.Template1;

    /// <summary>Emit the created table of contents.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = Document ?? WordDslContext.Require(this).Document;
        if (document == null)
        {
            throw new InvalidOperationException("Word document was not provided.");
        }

        var toc = document.AddTableOfContent(Style);

        if (PassThru.IsPresent)
        {
            WriteObject(toc);
        }
    }
}
