using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a single list item.</summary>
/// <para>Must be called within <c>Add-OfficeWordList</c>/<c>WordList</c>; supports nesting via <c>-Level</c>.</para>
/// <example>
///   <summary>Add bullet text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>WordList { Add-OfficeWordListItem -Text 'First task' }</code>
///   <para>Creates a bullet with the text “First task”.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordListItem")]
[Alias("WordListItem")]
public sealed class AddOfficeWordListItemCommand : PSCmdlet
{
    /// <summary>List item text.</summary>
    [Parameter(Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Zero-based list level.</summary>
    [Parameter]
    public int Level { get; set; }

    /// <summary>Emit the created <see cref="WordParagraph"/>.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = WordDslContext.Require(this);
        var list = context.CurrentList ?? throw new InvalidOperationException("WordListItem must be used inside WordList.");
        var anchor = context.ConsumeListAnchor(list);
        var paragraph = list.AddItem(string.IsNullOrEmpty(Text) ? null : Text, Level, anchor);

        if (anchor != null && !ReferenceEquals(paragraph, anchor) && string.IsNullOrWhiteSpace(anchor.Text))
        {
            anchor.Remove();
        }

        if (PassThru.IsPresent)
        {
            WriteObject(paragraph);
        }
    }
}
