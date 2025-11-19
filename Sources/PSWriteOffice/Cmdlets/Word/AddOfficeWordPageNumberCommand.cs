using System;
using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a PAGE field to the current header/footer.</summary>
/// <para>Typically used inside <c>WordFooter</c> to render “Page X of Y”.</para>
/// <example>
///   <summary>Footer page numbering.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>WordFooter { Add-OfficeWordPageNumber -IncludeTotalPages }</code>
///   <para>Outputs “Page # of #” in the footer.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordPageNumber")]
[Alias("WordPageNumber")]
public sealed class AddOfficeWordPageNumberCommand : PSCmdlet
{
    [Parameter]
    public SwitchParameter IncludeTotalPages { get; set; }

    [Parameter]
    public WordFieldFormat? Format { get; set; }

    [Parameter]
    public string Separator { get; set; } = " of ";

    protected override void ProcessRecord()
    {
        var context = WordDslContext.Require(this);
        WordHeaderFooter? target = context.CurrentFooter ?? context.CurrentHeader as WordHeaderFooter;
        if (target == null)
        {
            throw new InvalidOperationException("WordPageNumber must be called within WordHeader or WordFooter.");
        }

        target.AddPageNumber(IncludeTotalPages.IsPresent, Format, Separator);
    }
}
