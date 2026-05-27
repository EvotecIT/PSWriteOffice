using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Gets YAML front matter entries from a Markdown document.</summary>
/// <example>
///   <summary>Read Markdown front matter.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeMarkdownFrontMatter -Text "---`ntitle: Report`n---`n# Report"</code>
///   <para>Parses Markdown text and returns front matter entries.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeMarkdownFrontMatter", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(FrontMatterBlock.Entry))]
public sealed class GetOfficeMarkdownFrontMatterCommand : PSCmdlet
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

    /// <summary>Optional wildcard pattern matched against front matter keys.</summary>
    [Parameter]
    public string? Key { get; set; }

    /// <summary>Use case-sensitive matching for key filters.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
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
        var keyPattern = string.IsNullOrWhiteSpace(Key)
            ? null
            : new WildcardPattern(Key, wildcardOptions);

        foreach (var entry in document.FrontMatterEntries)
        {
            if (keyPattern != null && !keyPattern.IsMatch(entry.Key))
            {
                continue;
            }

            WriteObject(entry);
        }
    }
}
