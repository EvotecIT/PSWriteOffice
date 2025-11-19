using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds content to a section footer.</summary>
/// <para>Creates or reuses the requested footer part (default/first/even) and runs the DSL script block inside it.</para>
/// <example>
///   <summary>Append page numbers to the default footer.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordSection { Add-OfficeWordFooter { Add-OfficeWordPageNumber -IncludeTotalPages } }</code>
///   <para>Inserts a footer displaying “Page X of Y”.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordFooter")]
[Alias("WordFooter")]
public sealed class AddOfficeWordFooterCommand : PSCmdlet
{
    [Parameter]
    public HeaderFooterValues Type { get; set; } = HeaderFooterValues.Default;

    [Parameter]
    public ScriptBlock? Content { get; set; }

    protected override void ProcessRecord()
    {
        var context = WordDslContext.Require(this);
        var section = context.RequireSection();
        var footer = section.GetOrCreateFooter(Type);

        using (context.Push(footer))
        {
            Content?.InvokeReturnAsIs();
        }
    }
}
