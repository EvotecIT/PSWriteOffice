using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Services.Pdf;

internal static class PdfCommandUtilities
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

    internal static void EnsureDirectory(string path)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    internal static void EnsureOutputDirectory(string directory)
    {
        if (!Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    internal static string GetSafeFileName(string fileName)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var safe = new string(fileName.Select(character => invalid.Contains(character) ? '_' : character).ToArray());
        return string.IsNullOrWhiteSpace(safe) ? "attachment.bin" : safe;
    }

    internal static string GetUniquePath(string directory, string fileName)
    {
        var safeName = GetSafeFileName(fileName);
        var path = Path.Combine(directory, safeName);
        if (!File.Exists(path))
        {
            return path;
        }

        var extension = Path.GetExtension(safeName);
        var stem = Path.GetFileNameWithoutExtension(safeName);
        for (var index = 2; ; index++)
        {
            path = Path.Combine(directory, $"{stem}-{index}{extension}");
            if (!File.Exists(path))
            {
                return path;
            }
        }
    }

    internal static PdfDocument ResolveDocument(PSCmdlet cmdlet, PdfDocument? document, string parameterSetName, string documentParameterSet)
    {
        return parameterSetName == documentParameterSet
            ? document ?? throw new PSArgumentNullException(nameof(document))
            : PdfDslContext.Require(cmdlet).Document;
    }

    internal static PdfColor? ParseColor(string? color)
    {
        if (string.IsNullOrWhiteSpace(color))
        {
            return null;
        }

        var value = color!.Trim();
        if (value.StartsWith("#", StringComparison.Ordinal))
        {
            value = value.Substring(1);
        }

        if (value.Length != 6)
        {
            throw new PSArgumentException("Color must use #RRGGBB format.", nameof(color));
        }

        return PdfColor.FromRgb(
            byte.Parse(value.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
            byte.Parse(value.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
            byte.Parse(value.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
    }

    internal static PageSize ResolvePageSize(string? pageSize, double? width, double? height, bool landscape)
    {
        PageSize size = (pageSize ?? "Letter").Trim().ToUpperInvariant() switch
        {
            "A4" => PageSizes.A4,
            "A5" => PageSizes.A5,
            "LEGAL" => PageSizes.Legal,
            "LETTER" => PageSizes.Letter,
            "CUSTOM" => width.HasValue && height.HasValue
                ? new PageSize(width.Value, height.Value)
                : throw new PSArgumentException("Custom page size requires -Width and -Height."),
            _ => throw new PSArgumentException("PageSize must be A4, A5, Letter, Legal, or Custom.", nameof(pageSize))
        };

        return landscape ? size.Landscape() : size.Portrait();
    }

    internal static string[][] ConvertToTableRows(object[] inputObject, string[]? property, string[]? header)
    {
        if (inputObject.Length == 0)
        {
            throw new PSArgumentException("Provide at least one input object.", nameof(inputObject));
        }

        var propertyNames = property != null && property.Length > 0
            ? property
            : GetPropertyNames(inputObject[0]);

        var rows = new List<string[]>();
        if (header != null && header.Length > 0)
        {
            rows.Add(header);
        }
        else
        {
            rows.Add(propertyNames);
        }

        foreach (var item in inputObject)
        {
            if (item is IDictionary dictionary)
            {
                rows.Add(propertyNames.Select(name => TryGetDictionaryValue(dictionary, name, out var value)
                    ? Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty
                    : string.Empty).ToArray());
                continue;
            }

            var psObject = PSObject.AsPSObject(item);
            rows.Add(propertyNames.Select(name => Convert.ToString(psObject.Properties[name]?.Value, CultureInfo.InvariantCulture) ?? string.Empty).ToArray());
        }

        return rows.ToArray();
    }

    internal static string[][] ConvertDataRows(IEnumerable rows)
    {
        var result = new List<string[]>();
        foreach (var row in rows)
        {
            if (row is string[] strings)
            {
                result.Add(strings);
                continue;
            }

            if (row is IEnumerable enumerable && row is not string)
            {
                result.Add(enumerable.Cast<object?>().Select(value => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty).ToArray());
                continue;
            }

            result.Add(new[] { Convert.ToString(row, CultureInfo.InvariantCulture) ?? string.Empty });
        }

        if (result.Count == 0)
        {
            throw new PSArgumentException("Provide at least one table row.", nameof(rows));
        }

        return result.ToArray();
    }

    internal static IReadOnlyDictionary<string, string> ConvertFieldValues(IDictionary fieldValues)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (DictionaryEntry entry in fieldValues)
        {
            var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(key))
            {
                throw new PSArgumentException("Form field names cannot be empty.", nameof(fieldValues));
            }

            result[key!] = Convert.ToString(entry.Value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        return result;
    }

    internal static void ApplyPageRange(PdfTextStampOptions options, string? pageRange)
    {
        if (!string.IsNullOrWhiteSpace(pageRange))
        {
            options.UsePageRanges(PdfPageRange.ParseMany(pageRange!));
        }
    }

    internal static void ApplyPageRange(PdfImageStampOptions options, string? pageRange)
    {
        if (!string.IsNullOrWhiteSpace(pageRange))
        {
            options.UsePageRanges(PdfPageRange.ParseMany(pageRange!));
        }
    }

    private static string[] GetPropertyNames(object item)
    {
        if (item is IDictionary dictionary)
        {
            return dictionary.Keys.Cast<object>().Select(key => Convert.ToString(key, CultureInfo.InvariantCulture) ?? string.Empty).ToArray();
        }

        return PSObject.AsPSObject(item).Properties
            .Where(property => property.IsGettable)
            .Select(property => property.Name)
            .ToArray();
    }

    private static bool TryGetDictionaryValue(IDictionary dictionary, string key, out object? value)
    {
        if (dictionary.Contains(key))
        {
            value = dictionary[key];
            return true;
        }

        foreach (DictionaryEntry entry in dictionary)
        {
            if (string.Equals(Convert.ToString(entry.Key, CultureInfo.InvariantCulture), key, StringComparison.OrdinalIgnoreCase))
            {
                value = entry.Value;
                return true;
            }
        }

        value = null;
        return false;
    }
}
