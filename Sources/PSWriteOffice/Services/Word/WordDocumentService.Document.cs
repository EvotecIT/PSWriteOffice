using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

/// <summary>Bridges DSL cmdlets with OfficeIMO.Word document operations.</summary>
public static partial class WordDocumentService
{
    private static readonly FieldInfo? DisposedField = typeof(WordDocument).GetField("_disposed", BindingFlags.Instance | BindingFlags.NonPublic);
    private static readonly AsyncLocal<WordDocument[]?> TrackedDocuments = new();

    /// <summary>Loads an existing Word document.</summary>
    public static WordDocument LoadDocument(string filePath, bool readOnly, bool autoSave, string? password = null)
    {
        var resolvedPath = Path.GetFullPath(filePath);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"File {resolvedPath} doesn't exist.", resolvedPath);
        }

        if (!string.IsNullOrEmpty(password))
        {
            return RegisterDocument(OfficeEncryptedPackageService.LoadWord(resolvedPath, password!, readOnly, autoSave));
        }

        return RegisterDocument(WordDocument.Load(resolvedPath, readOnly, autoSave));
    }

    /// <summary>Creates a new Word document at the specified path.</summary>
    public static WordDocument CreateDocument(string filePath, bool autoSave)
    {
        return RegisterDocument(WordDocument.Create(Path.GetFullPath(filePath), autoSave));
    }

    /// <summary>Creates a new in-memory Word document without creating a package on disk.</summary>
    public static WordDocument CreateInMemoryDocument()
    {
        return RegisterDocument(WordDocument.Create());
    }

    /// <summary>Disposes the Word document.</summary>
    public static void CloseDocument(WordDocument document)
    {
        if (document == null)
        {
            throw new ArgumentNullException(nameof(document));
        }

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
        finally
        {
            UnregisterDocument(document);
        }
    }

    /// <summary>Saves the document, optionally to a new path, and closes it.</summary>
    public static void SaveDocument(WordDocument document, bool show, string? filePath, string? password = null)
    {
        if (string.IsNullOrWhiteSpace(document.FilePath) && string.IsNullOrWhiteSpace(filePath))
        {
            throw new InvalidOperationException("No file path provided.");
        }

        if (filePath != null)
        {
            SaveDocumentToPath(document, Path.GetFullPath(filePath), false, password);
        }
        else if (!string.IsNullOrEmpty(password))
        {
            var targetPath = document.FilePath!;
            OfficeEncryptedPackageService.SaveWord(document, targetPath, password!, false);
        }
        else
        {
            document.Save(false);
        }

        var savedPath = document.FilePath ?? filePath ?? throw new InvalidOperationException("No saved file path was available.");
        CloseDocument(document);

        if (show)
        {
            FileOpenService.Open(savedPath);
        }
    }

    private static void SaveDocumentToPath(WordDocument document, string path, bool openWord, string? password)
    {
        if (!string.IsNullOrEmpty(password))
        {
            OfficeEncryptedPackageService.SaveWord(document, path, password!, openWord);
            return;
        }

        document.Save(path, openWord);
    }

    /// <summary>Returns the most recently tracked Word document for the current runspace.</summary>
    public static WordDocument? GetCurrentTrackedDocument()
    {
        var tracked = GetAliveTrackedDocuments();
        return tracked.Length == 0
            ? null
            : tracked[tracked.Length - 1];
    }

    /// <summary>Returns tracked Word documents for the current runspace.</summary>
    public static IReadOnlyList<WordDocument> GetTrackedDocuments()
    {
        return GetAliveTrackedDocuments();
    }

    private static T RegisterDocument<T>(T document) where T : WordDocument
    {
        var tracked = GetAliveTrackedDocuments();
        TrackedDocuments.Value = tracked
            .Where(existing => !ReferenceEquals(existing, document))
            .Append(document)
            .ToArray();
        return document;
    }

    private static void UnregisterDocument(WordDocument document)
    {
        var tracked = TrackedDocuments.Value;
        if (tracked == null || tracked.Length == 0)
        {
            return;
        }

        TrackedDocuments.Value = tracked
            .Where(existing => !ReferenceEquals(existing, document))
            .ToArray();
    }

    private static WordDocument[] GetAliveTrackedDocuments()
    {
        var tracked = TrackedDocuments.Value ?? Array.Empty<WordDocument>();
        var alive = tracked.Where(document => !IsDisposed(document)).ToArray();
        if (alive.Length != tracked.Length)
        {
            TrackedDocuments.Value = alive;
        }

        return alive;
    }

    private static bool IsDisposed(WordDocument document)
    {
        if (DisposedField == null)
        {
            return false;
        }

        try
        {
            return (bool?)DisposedField.GetValue(document) ?? false;
        }
        catch
        {
            return false;
        }
    }
}
