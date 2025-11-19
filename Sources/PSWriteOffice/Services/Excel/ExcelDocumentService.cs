using System;
using System.IO;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelDocumentService
{
    public static ExcelDocument CreateDocument(string filePath, bool autoSave)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(filePath));
        }

        return ExcelDocument.Create(filePath);
    }

    public static ExcelDocument LoadDocument(string filePath, bool readOnly, bool autoSave)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"File '{filePath}' was not found.", filePath);
        }

        return ExcelDocument.Load(filePath, readOnly, autoSave);
    }

    public static void SaveDocument(ExcelDocument document, bool show, string? filePath)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));

        var currentPath = document.FilePath ?? string.Empty;
        if (!string.IsNullOrEmpty(filePath) && !string.Equals(filePath, currentPath, StringComparison.OrdinalIgnoreCase))
        {
            document.Save(filePath!, show);
        }
        else
        {
            document.Save(show);
        }

        document.Dispose();
    }

    public static void CloseDocument(ExcelDocument document)
    {
        document?.Dispose();
    }
}
