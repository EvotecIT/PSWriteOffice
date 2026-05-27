using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Updates styling on Word text runs returned by Get-OfficeWordRun.</summary>
/// <example>
///   <summary>Style runs selected from a document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeWordParagraph -Path .\Report.docx | Get-OfficeWordRun | Where-Object Text -eq 'Warning' | Set-OfficeWordRunStyle -Bold $true -Color '#C00000'</code>
///   <para>Applies bold red styling to matching runs.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeWordRunStyle")]
[Alias("WordRunStyle")]
[OutputType(typeof(WordParagraph))]
public sealed class SetOfficeWordRunStyleCommand : PSCmdlet
{
    /// <summary>Run to update. Runs are represented by OfficeIMO.Word.WordParagraph instances.</summary>
    [Parameter(ValueFromPipeline = true, Position = 0)]
    public WordParagraph? Run { get; set; }

    /// <summary>Replace the run text.</summary>
    [Parameter]
    public string? Text { get; set; }

    /// <summary>Character style to apply.</summary>
    [Parameter]
    public WordCharacterStyles? Style { get; set; }

    /// <summary>Character style id to apply.</summary>
    [Parameter]
    public string? StyleId { get; set; }

    /// <summary>Set or clear bold formatting.</summary>
    [Parameter]
    public bool? Bold { get; set; }

    /// <summary>Set or clear italic formatting.</summary>
    [Parameter]
    public bool? Italic { get; set; }

    /// <summary>Underline style to apply.</summary>
    [Parameter]
    public string? Underline { get; set; }

    /// <summary>Text color as #RRGGBB.</summary>
    [Parameter]
    public string? Color { get; set; }

    /// <summary>Font size in points.</summary>
    [Parameter]
    public int? FontSize { get; set; }

    /// <summary>Font family name.</summary>
    [Parameter]
    public string? FontFamily { get; set; }

    /// <summary>Highlight color.</summary>
    [Parameter]
    public string? Highlight { get; set; }

    /// <summary>Set or clear strikethrough.</summary>
    [Parameter]
    public bool? Strike { get; set; }

    /// <summary>Set or clear double strikethrough.</summary>
    [Parameter]
    public bool? DoubleStrike { get; set; }

    /// <summary>Capitalization style.</summary>
    [Parameter]
    public CapsStyle? CapsStyle { get; set; }

    /// <summary>Character spacing in twentieths of a point.</summary>
    [Parameter]
    public int? Spacing { get; set; }

    /// <summary>Set or clear outline effect.</summary>
    [Parameter]
    public bool? Outline { get; set; }

    /// <summary>Set or clear shadow effect.</summary>
    [Parameter]
    public bool? Shadow { get; set; }

    /// <summary>Set or clear emboss effect.</summary>
    [Parameter]
    public bool? Emboss { get; set; }

    /// <summary>Emit the updated run.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var run = Run ?? WordDslContext.Current?.CurrentParagraph;
        if (run == null)
        {
            throw new PSInvalidOperationException("Provide a Word run from Get-OfficeWordRun or run this command inside a Word paragraph DSL scope.");
        }

        if (IsBound(nameof(Text))) run.Text = Text ?? string.Empty;
        if (Style.HasValue) run.SetCharacterStyle(Style.Value);
        if (!string.IsNullOrWhiteSpace(StyleId)) run.SetCharacterStyleId(StyleId!);
        if (IsBound(nameof(Bold))) run.Bold = Bold ?? false;
        if (IsBound(nameof(Italic))) run.Italic = Italic ?? false;
        if (!string.IsNullOrWhiteSpace(Underline))
        {
            run.Underline = ParseOpenXmlValue<UnderlineValues>(Underline!, nameof(Underline));
        }
        if (IsBound(nameof(Color))) run.ColorHex = Color ?? string.Empty;
        if (IsBound(nameof(FontSize))) run.FontSize = FontSize;
        if (IsBound(nameof(FontFamily))) run.FontFamily = FontFamily;
        if (!string.IsNullOrWhiteSpace(Highlight))
        {
            run.Highlight = ParseOpenXmlValue<HighlightColorValues>(Highlight!, nameof(Highlight));
        }
        if (IsBound(nameof(Strike))) run.Strike = Strike ?? false;
        if (IsBound(nameof(DoubleStrike))) run.DoubleStrike = DoubleStrike ?? false;
        if (CapsStyle.HasValue) run.CapsStyle = CapsStyle.Value;
        if (IsBound(nameof(Spacing))) run.Spacing = Spacing;
        if (IsBound(nameof(Outline))) run.Outline = Outline ?? false;
        if (IsBound(nameof(Shadow))) run.Shadow = Shadow ?? false;
        if (IsBound(nameof(Emboss))) run.Emboss = Emboss ?? false;

        if (PassThru.IsPresent)
        {
            WriteObject(run);
        }
    }

    private static T ParseOpenXmlValue<T>(string value, string parameterName)
    {
        if (OpenXmlValueParser.TryParse(value, out T parsed))
        {
            return parsed;
        }

        throw new PSArgumentException($"Unknown {parameterName} value '{value}'.", parameterName);
    }

    private bool IsBound(string parameterName) => MyInvocation.BoundParameters.ContainsKey(parameterName);
}
