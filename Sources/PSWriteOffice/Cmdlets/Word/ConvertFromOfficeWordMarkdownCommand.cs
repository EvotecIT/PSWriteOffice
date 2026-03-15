using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates a Word document from Markdown.</summary>
/// <para>Returns a <see cref="WordDocument"/> or saves it to disk when <c>-OutputPath</c> is provided.</para>
/// <example>
///   <summary>Create a .docx from Markdown text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertFrom-OfficeWordMarkdown -Markdown '# Hello' -OutputPath .\hello.docx</code>
///   <para>Writes a Word document containing the supplied Markdown.</para>
/// </example>
/// <example>
///   <summary>Pipe a Markdown document into Word conversion.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeMarkdown -Path .\README.md | ConvertFrom-OfficeWordMarkdown</code>
///   <para>Returns a Word document instance for further edits.</para>
/// </example>
[Cmdlet(VerbsData.ConvertFrom, "OfficeWordMarkdown", DefaultParameterSetName = ParameterSetMarkdown)]
[Alias("ConvertFrom-WordMarkdown")]
[OutputType(typeof(WordDocument), typeof(FileInfo))]
public sealed class ConvertFromOfficeWordMarkdownCommand : PSCmdlet
{
    private const string ParameterSetMarkdown = "Markdown";
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown text to convert.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetMarkdown)]
    public string Markdown { get; set; } = string.Empty;

    /// <summary>Path to a Markdown file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("Path")]
    public string FilePath { get; set; } = string.Empty;

    /// <summary>Markdown document instance to convert.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Optional output path for the .docx file.</summary>
    [Parameter]
    [Alias("OutPath")]
    public string? OutputPath { get; set; }

    /// <summary>Optional font family applied during conversion.</summary>
    [Parameter]
    public string? FontFamily { get; set; }

    /// <summary>Base URI used to resolve relative links and images.</summary>
    [Parameter]
    [Alias("BasePath")]
    public string? BaseUri { get; set; }

    /// <summary>Allow local Markdown images to be inserted into the document.</summary>
    [Parameter]
    public SwitchParameter AllowLocalImages { get; set; }

    /// <summary>Restrict local images to one or more directories.</summary>
    [Parameter]
    public string[]? AllowedImageDirectory { get; set; }

    /// <summary>Allow remote HTTP(S) images to be downloaded and inserted.</summary>
    [Parameter]
    public SwitchParameter AllowRemoteImages { get; set; }

    /// <summary>Optional Markdown reader options used before Word conversion.</summary>
    [Parameter]
    public MarkdownReaderOptions? ReaderOptions { get; set; }

    /// <summary>Fit Markdown images to the page content width.</summary>
    [Parameter]
    public SwitchParameter FitImagesToPageContentWidth { get; set; }

    /// <summary>Fit Markdown images to the current content context width.</summary>
    [Parameter]
    public SwitchParameter FitImagesToContextWidth { get; set; }

    /// <summary>Optional hard cap for Markdown image width in pixels.</summary>
    [Parameter]
    public double? MaxImageWidthPixels { get; set; }

    /// <summary>Optional hard cap for Markdown image height in pixels.</summary>
    [Parameter]
    public double? MaxImageHeightPixels { get; set; }

    /// <summary>Optional hard cap for Markdown image width as a percentage of available content width.</summary>
    [Parameter]
    public double? MaxImageWidthPercentOfContent { get; set; }

    /// <summary>Open the document after saving.</summary>
    [Parameter]
    public SwitchParameter Open { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> when saving to disk.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WordDocument? document = null;

        try
        {
            if (FitImagesToPageContentWidth.IsPresent && FitImagesToContextWidth.IsPresent)
            {
                throw new ArgumentException("Use either -FitImagesToPageContentWidth or -FitImagesToContextWidth, not both.");
            }

            var options = BuildOptions();

            switch (ParameterSetName)
            {
                case ParameterSetPath:
                {
                    var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(FilePath);
                    if (!File.Exists(resolvedPath))
                    {
                        throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
                    }

                    if (string.IsNullOrWhiteSpace(options.BaseUri))
                    {
                        options.BaseUri = BuildDirectoryUri(Path.GetDirectoryName(resolvedPath) ?? Directory.GetCurrentDirectory());
                    }

                    document = WordMarkdownConverterExtensions.LoadFromMarkdown(resolvedPath, options);
                    break;
                }
                case ParameterSetDocument:
                    document = Document.ToWordDocument(options);
                    break;
                default:
                    if (string.IsNullOrWhiteSpace(Markdown))
                    {
                        throw new ArgumentException("Markdown content cannot be empty.", nameof(Markdown));
                    }

                    document = Markdown.LoadFromMarkdown(options);
                    break;
            }

            if (document == null)
            {
                throw new InvalidOperationException("Word document could not be created from Markdown.");
            }

            if (!string.IsNullOrWhiteSpace(OutputPath))
            {
                var resolvedOutput = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
                var directory = Path.GetDirectoryName(resolvedOutput);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                try
                {
                    document.Save(resolvedOutput, Open.IsPresent);
                }
                finally
                {
                    document.Dispose();
                }

                if (PassThru.IsPresent)
                {
                    WriteObject(new FileInfo(resolvedOutput));
                }
            }
            else
            {
                WriteObject(document);
            }
        }
        catch (Exception ex)
        {
            object? target = ParameterSetName switch
            {
                ParameterSetPath => FilePath,
                ParameterSetDocument => Document,
                _ => Markdown
            };
            WriteError(new ErrorRecord(ex, "MarkdownToWordFailed", ErrorCategory.InvalidOperation, target));
        }
    }

    private MarkdownToWordOptions BuildOptions()
    {
        var options = new MarkdownToWordOptions
        {
            AllowLocalImages = AllowLocalImages.IsPresent,
            AllowRemoteImages = AllowRemoteImages.IsPresent
        };

        if (!string.IsNullOrWhiteSpace(FontFamily))
        {
            options.FontFamily = FontFamily;
        }

        if (!string.IsNullOrWhiteSpace(BaseUri))
        {
            options.BaseUri = ResolveBaseUri(BaseUri!);
        }

        if (ReaderOptions != null)
        {
            options.ReaderOptions = ReaderOptions;
        }

        if (FitImagesToPageContentWidth.IsPresent)
        {
            options.FitImagesToPageContentWidth = true;
        }

        if (FitImagesToContextWidth.IsPresent)
        {
            options.FitImagesToContextWidth = true;
        }

        if (MaxImageWidthPixels.HasValue)
        {
            options.MaxImageWidthPixels = MaxImageWidthPixels.Value;
        }

        if (MaxImageHeightPixels.HasValue)
        {
            options.MaxImageHeightPixels = MaxImageHeightPixels.Value;
        }

        if (MaxImageWidthPercentOfContent.HasValue)
        {
            options.MaxImageWidthPercentOfContent = MaxImageWidthPercentOfContent.Value;
        }

        if (AllowedImageDirectory != null)
        {
            foreach (var entry in AllowedImageDirectory)
            {
                if (string.IsNullOrWhiteSpace(entry))
                {
                    continue;
                }

                options.AllowedImageDirectories.Add(SessionState.Path.GetUnresolvedProviderPathFromPSPath(entry));
            }
        }

        return options;
    }

    private string ResolveBaseUri(string value)
    {
        if (Uri.TryCreate(value, UriKind.Absolute, out var uri))
        {
            return uri.ToString();
        }

        var resolved = SessionState.Path.GetUnresolvedProviderPathFromPSPath(value);
        if (Directory.Exists(resolved))
        {
            return BuildDirectoryUri(resolved);
        }

        if (File.Exists(resolved))
        {
            return new Uri(resolved).AbsoluteUri;
        }

        if (Path.HasExtension(resolved))
        {
            return new Uri(Path.GetFullPath(resolved)).AbsoluteUri;
        }

        return BuildDirectoryUri(resolved);
    }

    private static string BuildDirectoryUri(string path)
    {
        var fullPath = Path.GetFullPath(path);
        if (!fullPath.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
            && !fullPath.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal))
        {
            fullPath += Path.DirectorySeparatorChar;
        }

        return new Uri(fullPath).AbsoluteUri;
    }
}
