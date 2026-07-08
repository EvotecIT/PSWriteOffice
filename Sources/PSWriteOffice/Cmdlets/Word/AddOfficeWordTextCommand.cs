using System;
using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds inline text to the current paragraph.</summary>
/// <para>Supports bold/italic/underline and color tweaks for quick DSL composition.</para>
/// <example>
///   <summary>Append bold text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph { Add-OfficeWordText -Text 'Important: ' -Bold }</code>
///   <para>Writes “Important:” with bold formatting.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordText", DefaultParameterSetName = "Text")]
[Alias("WordText", "WordBold", "WordItalic")]
public sealed class AddOfficeWordTextCommand : PSCmdlet
{
    private const string ParameterSetText = "Text";
    private const string ParameterSetRun = "Run";

    /// <summary>Text segments to append.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetText)]
    public string[] Text { get; set; } = Array.Empty<string>();

    /// <summary>Rich text runs. Each run can be created with TextRun/WordTextRun or provided as a hashtable/object.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetRun)]
    [Alias("Runs")]
    public object[]? Run { get; set; }

    /// <summary>Apply bold formatting.</summary>
    [Parameter]
    public SwitchParameter Bold { get; set; }

    /// <summary>Apply italic formatting.</summary>
    [Parameter]
    public SwitchParameter Italic { get; set; }

    /// <summary>Optional underline style.</summary>
    [Parameter]
    public UnderlineValues? Underline { get; set; }

    /// <summary>Run color (#RRGGBB).</summary>
    [Parameter]
    public string? Color { get; set; }

    /// <summary>Render text with strikethrough.</summary>
    [Parameter]
    public SwitchParameter Strike { get; set; }

    /// <summary>Font size in points.</summary>
    [Parameter]
    public int? FontSize { get; set; }

    /// <summary>Font name or family.</summary>
    [Parameter]
    [Alias("Font", "FontFamily", "Typeface")]
    public string? FontName { get; set; }

    /// <inheritdoc />
    protected override void BeginProcessing()
    {
        var name = MyInvocation.InvocationName;
        if (string.Equals(name, "WordBold", StringComparison.OrdinalIgnoreCase))
        {
            Bold = true;
        }
        else if (string.Equals(name, "WordItalic", StringComparison.OrdinalIgnoreCase))
        {
            Italic = true;
        }
    }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = WordDslContext.Require(this);
        var paragraph = context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();

        if (ParameterSetName == ParameterSetRun)
        {
            WordTextRunService.ApplyRuns(paragraph, Run!);
            return;
        }

        foreach (var entry in Text)
        {
            WordTextRunService.AddText(
                paragraph,
                entry,
                Bold.IsPresent,
                Italic.IsPresent,
                Underline,
                Strike.IsPresent,
                Color,
                FontSize,
                FontName);
        }
    }
}
