using System;
using System.IO;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

public static partial class WordDocumentService
{
    public static WordDocument LoadDocument(string filePath, bool readOnly, bool autoSave)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"File {filePath} doesn't exist.", filePath);
        }

        return WordDocument.Load(filePath, readOnly, autoSave);
    }

    public static WordDocument CreateDocument(string filePath, bool autoSave)
    {
        return WordDocument.Create(filePath, autoSave);
    }

    public static void CloseDocument(WordDocument document)
    {
        try
        {
            document.Dispose();
        }
        catch (Exception ex)
        {
            if (ex.InnerException?.Message != "Memory stream is not expandable.")
            {
                throw;
            }
        }
    }

    public static void SaveDocument(WordDocument document, bool show, string? filePath)
    {
        if (document.FilePath == null && filePath == null)
        {
            throw new InvalidOperationException("No file path provided.");
        }

        if (filePath != null)
        {
            document.Save(filePath, show);
        }
        else if (!document.AutoSave)
        {
            document.Save(show);
        }

        document.Dispose();
    }
}
