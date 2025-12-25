using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds or reuses a section inside the current Word document.</summary>
/// <para>Provides the DSL entry point for section-level operations inside <c>New-OfficeWord</c>.</para>
/// <example>
///   <summary>Create a section with a paragraph.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\doc.docx { Add-OfficeWordSection { Add-OfficeWordParagraph -Text 'Hello' } }</code>
///   <para>Creates a document and inserts a section that contains a single paragraph.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordSection")]
[Alias("WordSection")]
public sealed class AddOfficeWordSectionCommand : PSCmdlet
{
    /// <summary>DSL scriptblock executed within the section scope.</summary>
    [Parameter(Position = 0)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Optional section break type.</summary>
    [Parameter]
    public SectionMarkValues? BreakType { get; set; }

    /// <summary>Emit the created <see cref="WordSection"/>.</summary>
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = WordDslContext.Require(this);
        var section = context.AcquireSection(BreakType);

        using (context.Push(section))
        {
            Content?.InvokeReturnAsIs();
        }

        if (PassThru.IsPresent)
        {
            WriteObject(section);
        }
    }
}
