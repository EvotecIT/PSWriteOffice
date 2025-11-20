using System;
using System.IO;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

/// <summary>Bridges DSL cmdlets with OfficeIMO.Word document operations.</summary>
public static partial class WordDocumentService
{
    /// <summary>Loads an existing Word document.</summary>
    public static WordDocument LoadDocument(string filePath, bool readOnly, bool autoSave)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"File {filePath} doesn't exist.", filePath);
        }

        return WordDocument.Load(filePath, readOnly, autoSave);
    }

    /// <summary>Creates a new Word document at the specified path.</summary>
    public static WordDocument CreateDocument(string filePath, bool autoSave)
    {
        return WordDocument.Create(filePath, autoSave);
    }

    /// <summary>Disposes the Word document.</summary>
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

    /// <summary>Saves the document, optionally to a new path, and closes it.</summary>
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
        else
        {
            document.Save(show);
        }

        document.Dispose();
    }
}
