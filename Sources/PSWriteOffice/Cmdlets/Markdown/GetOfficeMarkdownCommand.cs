using System.IO;
using System.Management.Automation;
using OfficeIMO.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Parses Markdown text or files into a Markdown document model.</summary>
/// <para>Returns an <see cref="MarkdownDoc"/> for inspection or further rendering.</para>
/// <example>
///   <summary>Parse a Markdown file.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$md = Get-OfficeMarkdown -Path .\README.md</code>
///   <para>Loads the file into a Markdown document object.</para>
/// </example>
/// <example>
///   <summary>Parse Markdown text in-memory.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$md = Get-OfficeMarkdown -Text '# Title`n`nBody text'</code>
///   <para>Parses Markdown text directly into a document model.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeMarkdown", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(MarkdownDoc))]
public sealed class GetOfficeMarkdownCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetText = "Text";

    /// <summary>Path to the Markdown file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath", "Path")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Markdown text to parse.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetText)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Optional reader options.</summary>
    [Parameter]
    public MarkdownReaderOptions? Options { get; set; }

    /// <summary>Named reader profile used when <see cref="Options"/> is not supplied.</summary>
    [Parameter]
    public MarkdownReaderOptions.MarkdownDialectProfile? Profile { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Options != null && Profile.HasValue)
        {
            throw new PSArgumentException("Specify either -Options or -Profile, not both.");
        }

        var options = Options ?? (Profile.HasValue
            ? MarkdownReaderOptions.CreateProfile(Profile.Value)
            : null);

        MarkdownDoc document;
        if (ParameterSetName == ParameterSetPath)
        {
            var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
            if (!File.Exists(resolvedPath))
            {
                throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
            }

            document = MarkdownReader.ParseFile(resolvedPath, options);
        }
        else
        {
            document = MarkdownReader.Parse(Text ?? string.Empty, options);
        }

        WriteObject(document);
    }
}
