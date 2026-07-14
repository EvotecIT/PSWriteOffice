using System;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace PSWriteOffice.Services;

internal static class OfficeEncryptedPackageService
{
    public static ExcelDocument LoadExcel(string path, string password, bool readOnly, bool autoSave)
    {
        if (autoSave)
        {
            throw new NotSupportedException("Encrypted Excel workbooks require explicit Save-OfficeExcel -Password or Close-OfficeExcel -Save -Password. OfficeIMO does not support SaveOnDispose for encrypted sources.");
        }

        return ExcelDocument.LoadEncrypted(path, password, new ExcelLoadOptions
        {
            AccessMode = readOnly ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite,
            PersistenceMode = DocumentPersistenceMode.Explicit
        });
    }

    public static void SaveExcel(ExcelDocument document, string path, string password, bool openExcel, ExcelSaveOptions? saveOptions)
    {
        document.SaveEncrypted(path, password, saveOptions);
        if (openExcel)
        {
            FileOpenService.Open(path);
        }
    }

    public static WordDocument LoadWord(string path, string password, bool readOnly, bool autoSave)
    {
        if (autoSave)
        {
            throw new NotSupportedException("Encrypted Word documents require explicit Save-OfficeWord -Password or Close-OfficeWord -Save -Password. OfficeIMO does not support SaveOnDispose for encrypted sources.");
        }

        return WordDocument.LoadEncrypted(path, password, new WordLoadOptions
        {
            AccessMode = readOnly ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite,
            PersistenceMode = DocumentPersistenceMode.Explicit
        });
    }

    public static void SaveWord(WordDocument document, string path, string password, bool openWord)
    {
        document.SaveEncrypted(path, password);
        if (openWord)
        {
            FileOpenService.Open(path);
        }
    }

    public static PowerPointPresentation OpenPowerPoint(string path, string password)
    {
        return PowerPointPresentation.LoadEncrypted(path, password, new PowerPointLoadOptions
        {
            AccessMode = DocumentAccessMode.ReadWrite,
            PersistenceMode = DocumentPersistenceMode.Explicit
        });
    }

    public static void SavePowerPoint(PowerPointPresentation presentation, Stream stream, string password)
    {
        presentation.SaveEncrypted(stream, password);
    }
}
