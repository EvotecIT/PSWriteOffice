using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Gets heading metadata from a Markdown document.</summary>
/// <para>Returns heading level, text, resolved anchor, and the backing heading block.</para>
/// <example>
///   <summary>Read headings from Markdown text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeMarkdownHeading -Text "# Title`n`n## Details"</code>
///   <para>Parses Markdown text and returns the document headings.</para>
/// </example>
/// <example>
///   <summary>Inspect a parsed Markdown document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeMarkdown -Path .\README.md | Get-OfficeMarkdownHeading -MinLevel 2</code>
///   <para>Returns headings from an existing Markdown document object.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeMarkdownHeading", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(MarkdownDoc.HeadingInfo))]
public sealed class GetOfficeMarkdownHeadingCommand : PSCmdlet
{
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";
    private const string ParameterSetText = "Text";

    /// <summary>Markdown document to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Path to the Markdown file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath", "Path")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Markdown text to parse.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetText)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Optional reader options used when parsing path or text input.</summary>
    [Parameter]
    public MarkdownReaderOptions? Options { get; set; }

    /// <summary>Named reader profile used when <see cref="Options"/> is not supplied.</summary>
    [Parameter]
    public MarkdownReaderOptions.MarkdownDialectProfile? Profile { get; set; }

    /// <summary>Minimum heading level to return.</summary>
    [Parameter]
    [ValidateRange(1, 6)]
    public int MinLevel { get; set; } = 1;

    /// <summary>Maximum heading level to return.</summary>
    [Parameter]
    [ValidateRange(1, 6)]
    public int MaxLevel { get; set; } = 6;

    /// <summary>Optional wildcard pattern matched against heading text.</summary>
    [Parameter]
    public string? HeadingText { get; set; }

    /// <summary>Optional wildcard pattern matched against resolved heading anchors.</summary>
    [Parameter]
    public string? Anchor { get; set; }

    /// <summary>Use case-sensitive matching for text and anchor filters.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (MinLevel > MaxLevel)
        {
            throw new PSArgumentException("-MinLevel cannot be greater than -MaxLevel.");
        }

        var document = MarkdownDocumentResolver.Resolve(
            this,
            ParameterSetName,
            ParameterSetDocument,
            Document,
            InputPath,
            Text,
            Options,
            Profile);

        var wildcardOptions = CaseSensitive
            ? WildcardOptions.None
            : WildcardOptions.IgnoreCase;
        var textPattern = string.IsNullOrWhiteSpace(HeadingText)
            ? null
            : new WildcardPattern(HeadingText, wildcardOptions);
        var anchorPattern = string.IsNullOrWhiteSpace(Anchor)
            ? null
            : new WildcardPattern(Anchor!.TrimStart('#'), wildcardOptions);

        foreach (var heading in document.GetHeadingInfos())
        {
            if (heading.Level < MinLevel || heading.Level > MaxLevel)
            {
                continue;
            }

            if (textPattern != null && !textPattern.IsMatch(heading.Text))
            {
                continue;
            }

            if (anchorPattern != null && !anchorPattern.IsMatch(heading.Anchor))
            {
                continue;
            }

            WriteObject(heading);
        }
    }
}
