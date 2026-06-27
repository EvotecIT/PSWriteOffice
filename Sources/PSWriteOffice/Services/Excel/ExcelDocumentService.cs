using System;
using System.IO;
using OfficeIMO.Excel;
using PSWriteOffice.Services;

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

    public static ExcelDocument CreateInMemoryDocument()
    {
        return ExcelDocument.Create(new MemoryStream(), autoSave: false);
    }

    public static ExcelDocument CreateDocumentFromTemplate(string templatePath, string filePath, bool autoSave)
    {
        if (string.IsNullOrWhiteSpace(templatePath))
        {
            throw new ArgumentException("Template path cannot be empty.", nameof(templatePath));
        }

        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(filePath));
        }

        return ExcelDocument.CreateFromTemplate(
            Path.GetFullPath(templatePath),
            Path.GetFullPath(filePath),
            overwrite: true,
            autoSave: autoSave);
    }

    public static void CopyWorkbookPackage(string sourcePath, string destinationPath, bool overwrite)
    {
        if (string.IsNullOrWhiteSpace(sourcePath))
        {
            throw new ArgumentException("Source path cannot be empty.", nameof(sourcePath));
        }

        if (string.IsNullOrWhiteSpace(destinationPath))
        {
            throw new ArgumentException("Destination path cannot be empty.", nameof(destinationPath));
        }

        ExcelDocument.CopyPackage(
            Path.GetFullPath(sourcePath),
            Path.GetFullPath(destinationPath),
            overwrite: overwrite);
    }

    public static ExcelDocument LoadDocument(string filePath, bool readOnly, bool autoSave, string? password = null)
    {
        var resolvedPath = Path.GetFullPath(filePath);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
        }

        if (!string.IsNullOrEmpty(password))
        {
            return OfficeEncryptedPackageService.LoadExcel(resolvedPath, password!, readOnly, autoSave);
        }

        return ExcelDocument.Load(resolvedPath, readOnly, autoSave);
    }

    public static ExcelDocument LoadDocument(Uri uri, bool readOnly, bool allowHttp, string? password = null)
    {
        if (!string.IsNullOrEmpty(password))
        {
            throw new NotSupportedException("Encrypted remote workbook loads are not supported.");
        }

        return ExcelDocument.Load(uri, ExcelHttpLoadService.CreateOptions(allowHttp), readOnly);
    }

    public static void SaveDocument(ExcelDocument document, bool show, string? filePath, string? password = null, ExcelSaveOptions? saveOptions = null)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));

        var currentPath = document.FilePath ?? string.Empty;
        if (!string.IsNullOrEmpty(filePath))
        {
            var target = filePath!;
            if (!string.Equals(target, currentPath, StringComparison.OrdinalIgnoreCase))
            {
                SaveDocumentToPath(document, Path.GetFullPath(target), false, password, saveOptions);
                var savedAsPath = document.FilePath ?? target;
                document.Dispose();
                if (show)
                {
                    FileOpenService.Open(savedAsPath);
                }
                return;
            }
        }

        if (!string.IsNullOrEmpty(password))
        {
            if (string.IsNullOrWhiteSpace(document.FilePath))
            {
                throw new InvalidOperationException("No file path provided for encrypted save.");
            }

            var targetPath = document.FilePath!;
            OfficeEncryptedPackageService.SaveExcel(document, targetPath, password!, false, saveOptions);
        }
        else
        {
            if (saveOptions == null)
            {
                document.Save(false);
            }
            else if (!string.IsNullOrWhiteSpace(document.FilePath))
            {
                var targetPath = document.FilePath!;
                document.Save(targetPath, false, saveOptions);
            }
            else
            {
                throw new InvalidOperationException("No file path provided for save options.");
            }
        }

        var savedPath = document.FilePath ?? filePath ?? throw new InvalidOperationException("No saved file path was available.");
        document.Dispose();
        if (show)
        {
            FileOpenService.Open(savedPath);
        }
    }

    public static ExcelSaveOptions? CreateSaveOptions(
        bool safePreflight,
        bool safeRepairDefinedNames,
        bool validateOpenXml,
        bool disableFastPackageWriter = false,
        bool evaluateFormulas = false,
        bool clearCachedFormulaResults = false,
        bool markFormulasDirty = false,
        bool forceFullCalculationOnOpen = false)
    {
        if (!safePreflight &&
            !safeRepairDefinedNames &&
            !validateOpenXml &&
            !disableFastPackageWriter &&
            !evaluateFormulas &&
            !clearCachedFormulaResults &&
            !markFormulasDirty &&
            !forceFullCalculationOnOpen)
        {
            return null;
        }

        return new ExcelSaveOptions
        {
            SafePreflight = safePreflight,
            SafeRepairDefinedNames = safeRepairDefinedNames,
            ValidateOpenXml = validateOpenXml,
            DisableFastPackageWriter = disableFastPackageWriter,
            EvaluateFormulasBeforeSave = evaluateFormulas,
            ClearCachedFormulaResultsBeforeSave = clearCachedFormulaResults,
            MarkFormulasDirtyBeforeSave = markFormulasDirty,
            ForceFullCalculationOnOpen = forceFullCalculationOnOpen
        };
    }

    private static void SaveDocumentToPath(ExcelDocument document, string path, bool openExcel, string? password, ExcelSaveOptions? saveOptions)
    {
        if (!string.IsNullOrEmpty(password))
        {
            OfficeEncryptedPackageService.SaveExcel(document, path, password!, openExcel, saveOptions);
            return;
        }

        document.Save(path, openExcel, saveOptions);
    }

    public static void CloseDocument(ExcelDocument document)
    {
        document?.Dispose();
    }
}
