using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Converts Markdown content to HTML.</summary>
/// <para>Returns HTML text or saves it to a file when <c>-OutputPath</c> is specified.</para>
/// <example>
///   <summary>Convert a Markdown file to HTML.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$html = ConvertTo-OfficeMarkdownHtml -Path .\README.md</code>
///   <para>Returns the rendered HTML.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeMarkdownHtml", DefaultParameterSetName = ParameterSetPath)]
[Alias("Convert-MarkdownToHtml")]
[OutputType(typeof(string), typeof(FileInfo))]
public sealed class ConvertToOfficeMarkdownHtmlCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetText = "Text";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the Markdown file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath", "Path")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Markdown text to convert.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetText)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Markdown document to convert.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Optional output path for the HTML file.</summary>
    [Parameter]
    [Alias("OutPath")]
    public string? OutputPath { get; set; }

    /// <summary>Render a full HTML document instead of a fragment.</summary>
    [Parameter]
    public SwitchParameter DocumentMode { get; set; }

    /// <summary>Built-in HTML style preset.</summary>
    [Parameter]
    public HtmlStyle Style { get; set; } = HtmlStyle.Clean;

    /// <summary>CSS delivery mode.</summary>
    [Parameter]
    public CssDelivery CssDelivery { get; set; } = CssDelivery.Inline;

    /// <summary>Asset loading mode.</summary>
    [Parameter]
    public AssetMode AssetMode { get; set; } = AssetMode.Online;

    /// <summary>Optional title for HTML documents.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Optional reader options when parsing Markdown.</summary>
    [Parameter]
    public MarkdownReaderOptions? ReaderOptions { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> when saving to disk.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        MarkdownDoc document;
        if (ParameterSetName == ParameterSetDocument)
        {
            document = Document ?? throw new InvalidOperationException("Markdown document was not provided.");
        }
        else if (ParameterSetName == ParameterSetPath)
        {
            var resolved = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
            if (!File.Exists(resolved))
            {
                throw new FileNotFoundException($"File '{resolved}' was not found.", resolved);
            }
            document = MarkdownReader.ParseFile(resolved, ReaderOptions);
        }
        else
        {
            document = MarkdownReader.Parse(Text ?? string.Empty, ReaderOptions);
        }

        var options = new HtmlOptions
        {
            Kind = DocumentMode.IsPresent ? HtmlKind.Document : HtmlKind.Fragment,
            Style = Style,
            CssDelivery = CssDelivery,
            AssetMode = AssetMode
        };

        if (!string.IsNullOrWhiteSpace(Title))
        {
            options.Title = Title;
        }

        if (!string.IsNullOrWhiteSpace(OutputPath))
        {
            var resolvedOutput = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
            var directory = Path.GetDirectoryName(resolvedOutput);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            document.SaveHtml(resolvedOutput, options);
            if (PassThru.IsPresent)
            {
                WriteObject(new FileInfo(resolvedOutput));
            }
        }
        else
        {
            WriteObject(options.Kind == HtmlKind.Document
                ? document.ToHtmlDocument(options)
                : document.ToHtmlFragment(options));
        }
    }
}
