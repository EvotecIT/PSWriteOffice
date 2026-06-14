using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Json;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Visio;
using OfficeIMO.Reader.Xml;
using OfficeIMO.Reader.Yaml;
using OfficeIMO.Reader.Zip;

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

    internal static void RegisterReaderAdapters()
    {
        RegisterPdfReader();

        RegisterAdapter(
            DocumentReaderHtmlRegistrationExtensions.HandlerId,
            new[] { ".html", ".htm", ".xhtml" },
            static () => DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler());
        RegisterAdapter(
            DocumentReaderCsvRegistrationExtensions.HandlerId,
            new[] { ".csv", ".tsv" },
            static () => DocumentReaderCsvRegistrationExtensions.RegisterCsvHandler(replaceExisting: true));
        RegisterAdapter(
            DocumentReaderJsonRegistrationExtensions.HandlerId,
            new[] { ".json" },
            static () => DocumentReaderJsonRegistrationExtensions.RegisterJsonHandler(replaceExisting: true));
        RegisterAdapter(
            DocumentReaderXmlRegistrationExtensions.HandlerId,
            new[] { ".xml" },
            static () => DocumentReaderXmlRegistrationExtensions.RegisterXmlHandler(replaceExisting: true));
        RegisterAdapter(
            DocumentReaderYamlRegistrationExtensions.HandlerId,
            new[] { ".yaml", ".yml" },
            static () => DocumentReaderYamlRegistrationExtensions.RegisterYamlHandler(replaceExisting: true));
        RegisterAdapter(
            DocumentReaderZipRegistrationExtensions.HandlerId,
            new[] { ".zip" },
            static () => DocumentReaderZipRegistrationExtensions.RegisterZipHandler());
        RegisterAdapter(
            DocumentReaderEpubRegistrationExtensions.HandlerId,
            new[] { ".epub" },
            static () => DocumentReaderEpubRegistrationExtensions.RegisterEpubHandler());
        RegisterAdapter(
            DocumentReaderVisioRegistrationExtensions.HandlerId,
            new[] { ".vsdx", ".vsdm", ".vstx", ".vstm" },
            static () => DocumentReaderVisioRegistrationExtensions.RegisterVisioHandler());
    }

    private static void RegisterAdapter(string handlerId, string[] extensions, Action register)
    {
        var customCapabilities = DocumentReader.GetCapabilities(includeBuiltIn: false, includeCustom: true);
        if (customCapabilities.Any(capability => string.Equals(capability.Id, handlerId, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        if (customCapabilities.Any(capability => capability.Extensions.Any(extension =>
            extensions.Any(target => string.Equals(extension, target, StringComparison.OrdinalIgnoreCase)))))
        {
            return;
        }

        register();
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
