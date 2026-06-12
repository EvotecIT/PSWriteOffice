using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;

namespace PSWriteOffice.Services.Reader;

internal static class ReaderCommandUtilities
{
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

    internal static void RegisterPdfReader()
    {
        var customPdfHandler = DocumentReader.GetCapabilities(includeBuiltIn: false, includeCustom: true)
            .FirstOrDefault(static capability => capability.Extensions.Any(static extension =>
                string.Equals(extension, ".pdf", StringComparison.OrdinalIgnoreCase)));

        if (customPdfHandler != null)
        {
            return;
        }

        DocumentReaderPdfRegistrationExtensions.RegisterPdfHandler(replaceExisting: true);
    }

    internal static ReaderOptions BuildReaderOptions(
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
        var options = new ReaderOptions
        {
            IncludeWordFootnotes = includeWordFootnotes,
            IncludePowerPointNotes = includePowerPointNotes,
            ExcelHeadersInFirstRow = excelHeadersInFirstRow,
            MarkdownChunkByHeadings = markdownChunkByHeadings,
            ComputeHashes = computeHashes
        };

        if (maxInputBytes.HasValue)
        {
            options.MaxInputBytes = maxInputBytes.Value;
        }

        if (openXmlMaxCharactersInPart.HasValue)
        {
            options.OpenXmlMaxCharactersInPart = openXmlMaxCharactersInPart.Value;
        }

        if (maxChars.HasValue)
        {
            options.MaxChars = maxChars.Value;
        }

        if (maxTableRows.HasValue)
        {
            options.MaxTableRows = maxTableRows.Value;
        }

        if (excelChunkRows.HasValue)
        {
            options.ExcelChunkRows = excelChunkRows.Value;
        }

        if (!string.IsNullOrWhiteSpace(excelSheetName))
        {
            options.ExcelSheetName = excelSheetName;
        }

        if (!string.IsNullOrWhiteSpace(excelA1Range))
        {
            options.ExcelA1Range = excelA1Range;
        }

        return options;
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
