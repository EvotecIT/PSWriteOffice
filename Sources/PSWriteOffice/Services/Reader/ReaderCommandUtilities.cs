using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;

namespace PSWriteOffice.Services.Reader;

internal static class ReaderCommandUtilities
{
    private static readonly Lazy<OfficeDocumentReader> SharedReader = new(() => CreateReader());

    internal static OfficeDocumentReader Reader => SharedReader.Value;

    internal static string ResolvePath(PSCmdlet cmdlet, string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new PSArgumentException("Path cannot be empty.", nameof(path));
        }

        var providerPath = cmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath(path);
        return Path.IsPathRooted(providerPath)
            ? providerPath
            : Path.Combine(cmdlet.SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }

    internal static OfficeDocumentReaderBuilder CreateBuilder(ReaderAllOptions? options = null)
    {
        return new OfficeDocumentReaderBuilder()
            .AddAllOfficeIMOHandlers(options);
    }

    internal static OfficeDocumentReader CreateReader(ReaderAllOptions? options = null) => CreateBuilder(options).Build();

    internal static ReaderCommandConfiguration BuildReadConfiguration(
        long? maxInputBytes,
        long? openXmlMaxCharactersInPart,
        int? maxChars,
        int? maxTableRows,
        bool includeWordFootnotes,
        bool includePowerPointNotes,
        bool excelHeadersInFirstRow,
        int? excelChunkRows,
        string? excelSheetName,
        string? excelA1Range,
        bool markdownChunkByHeadings,
        bool computeHashes)
    {
        var readerOptions = new ReaderOptions
        {
            ComputeHashes = computeHashes
        };

        if (maxInputBytes.HasValue)
        {
            readerOptions.MaxInputBytes = maxInputBytes.Value;
        }

        if (openXmlMaxCharactersInPart.HasValue)
        {
            readerOptions.OpenXmlMaxCharactersInPart = openXmlMaxCharactersInPart.Value;
        }

        if (maxChars.HasValue)
        {
            readerOptions.MaxChars = maxChars.Value;
        }

        if (maxTableRows.HasValue)
        {
            readerOptions.MaxTableRows = maxTableRows.Value;
        }

        var hasHandlerOverrides =
            !includeWordFootnotes ||
            !includePowerPointNotes ||
            !excelHeadersInFirstRow ||
            excelChunkRows.HasValue ||
            !string.IsNullOrWhiteSpace(excelSheetName) ||
            !string.IsNullOrWhiteSpace(excelA1Range) ||
            !markdownChunkByHeadings;
        ReaderAllOptions? handlerOptions = hasHandlerOverrides
            ? new ReaderAllOptions
            {
                Word = new OfficeIMO.Reader.Word.ReaderWordOptions
                {
                    IncludeFootnotes = includeWordFootnotes
                },
                PowerPoint = new OfficeIMO.Reader.PowerPoint.ReaderPowerPointOptions
                {
                    IncludeNotes = includePowerPointNotes
                },
                Excel = new OfficeIMO.Reader.Excel.ReaderExcelOptions
                {
                    HeadersInFirstRow = excelHeadersInFirstRow,
                    ChunkRows = excelChunkRows ?? 200,
                    SheetName = string.IsNullOrWhiteSpace(excelSheetName) ? null : excelSheetName,
                    A1Range = string.IsNullOrWhiteSpace(excelA1Range) ? null : excelA1Range
                },
                Markdown = new OfficeIMO.Reader.Markdown.ReaderMarkdownOptions
                {
                    ChunkByHeadings = markdownChunkByHeadings
                }
            }
            : null;

        return new ReaderCommandConfiguration(readerOptions, handlerOptions);
    }

    internal static ReaderFolderOptions BuildFolderOptions(bool recurse, int? maxFiles, long? maxTotalBytes, string[]? extension)
    {
        var options = new ReaderFolderOptions
        {
            Recurse = recurse
        };

        if (maxFiles.HasValue)
        {
            options.MaxFiles = maxFiles.Value;
        }

        if (maxTotalBytes.HasValue)
        {
            options.MaxTotalBytes = maxTotalBytes.Value;
        }

        if (extension != null && extension.Length > 0)
        {
            options.Extensions = extension
                .Where(static value => !string.IsNullOrWhiteSpace(value))
                .Select(static value => value.StartsWith(".", StringComparison.Ordinal) ? value : "." + value)
                .ToArray();
        }

        return options;
    }
}

internal sealed class ReaderCommandConfiguration
{
    internal ReaderCommandConfiguration(ReaderOptions readerOptions, ReaderAllOptions? handlerOptions)
    {
        ReaderOptions = readerOptions;
        HandlerOptions = handlerOptions;
    }

    internal ReaderOptions ReaderOptions { get; }

    internal ReaderAllOptions? HandlerOptions { get; }
}
