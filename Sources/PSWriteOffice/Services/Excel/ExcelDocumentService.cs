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

        return ExcelDocument.Create(Path.GetFullPath(filePath), autoSave);
    }

    public static ExcelDocument LoadDocument(string filePath, bool readOnly, bool autoSave)
    {
        var resolvedPath = Path.GetFullPath(filePath);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
        }

        return ExcelDocument.Load(resolvedPath, readOnly, autoSave);
    }

    public static void SaveDocument(ExcelDocument document, bool show, string? filePath)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));

        var currentPath = document.FilePath ?? string.Empty;
        if (!string.IsNullOrEmpty(filePath))
        {
            var target = filePath!;
            if (!string.Equals(target, currentPath, StringComparison.OrdinalIgnoreCase))
            {
                document.Save(Path.GetFullPath(target), false);
                var savedAsPath = document.FilePath ?? target;
                document.Dispose();
                if (show)
                {
                    FileOpenService.Open(savedAsPath);
                }
                return;
            }
        }

        document.Save(false);
        var savedPath = document.FilePath ?? filePath ?? throw new InvalidOperationException("No saved file path was available.");
        document.Dispose();
        if (show)
        {
            FileOpenService.Open(savedPath);
        }
    }

    public static void CloseDocument(ExcelDocument document)
    {
        document?.Dispose();
    }
}
