using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Removes the table of contents from a Word document.</summary>
/// <example>
///   <summary>Remove the table of contents.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Remove-OfficeWordTableOfContent</code>
///   <para>Deletes the table of contents if one exists.</para>
/// </example>
[Cmdlet(VerbsCommon.Remove, "OfficeWordTableOfContent")]
public sealed class RemoveOfficeWordTableOfContentCommand : PSCmdlet
{
    /// <summary>Document to modify when provided explicitly.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordDocument? Document { get; set; }

    /// <summary>Emit the document after removal.</summary>
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

        document.RemoveTableOfContent();

        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
