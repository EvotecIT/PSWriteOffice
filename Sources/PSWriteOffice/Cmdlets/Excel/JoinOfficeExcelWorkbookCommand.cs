using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Merges worksheets from one or more workbooks into a target workbook.</summary>
/// <example>
///   <summary>Merge workbook sheets using the package fast path.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$sources = Get-ChildItem .\Incoming\*.xlsx | Select-Object -ExpandProperty FullName
/// $results = Join-OfficeExcelWorkbook -Path .\Combined.xlsx -SourcePath $sources -CopyMode Package -SheetNamePrefix Import
/// $results | Select-Object SheetCount, SourceSheets, TargetSheets</code>
///   <para>Copies worksheets between packages without importing rows into PowerShell objects, which is the preferred path for large workbook merge workflows.</para>
/// </example>
[Cmdlet(VerbsCommon.Join, "OfficeExcelWorkbook")]
[Alias("Merge-OfficeExcelWorkbook", "ExcelWorkbookJoin", "ExcelWorkbookMerge")]
[OutputType(typeof(ExcelWorkbookMergeResult))]
public sealed class JoinOfficeExcelWorkbookCommand : PSCmdlet
{
    /// <summary>Target workbook path to create or update.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("Path", "FilePath", "OutputPath")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Source workbook paths to merge into the target workbook.</summary>
    [Parameter(Mandatory = true, Position = 1, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
    [Alias("FullName", "LiteralPath")]
    public string[] SourcePath { get; set; } = [];

    /// <summary>Optional source worksheet names. By default all worksheets are copied.</summary>
    [Parameter]
    [Alias("Sheet", "WorksheetName")]
    public string[]? SheetName { get; set; }

    /// <summary>Optional prefix applied to copied worksheet names.</summary>
    [Parameter]
    public string? SheetNamePrefix { get; set; }

    /// <summary>Controls how invalid or duplicate destination sheet names are handled.</summary>
    [Parameter]
    public SheetNameValidationMode ValidationMode { get; set; } = SheetNameValidationMode.Sanitize;

    /// <summary>Controls whether cross-workbook copies use package-level copy or value materialization.</summary>
    [Parameter]
    public ExcelWorksheetCopyMode CopyMode { get; set; } = ExcelWorksheetCopyMode.Package;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var resolvedTargetPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
        var targetDirectory = Path.GetDirectoryName(resolvedTargetPath);
        if (!string.IsNullOrWhiteSpace(targetDirectory))
        {
            Directory.CreateDirectory(targetDirectory);
        }

        using var targetDocument = File.Exists(resolvedTargetPath)
            ? ExcelDocumentService.LoadDocument(resolvedTargetPath, readOnly: false, autoSave: false)
            : ExcelDocument.Create(resolvedTargetPath, autoSave: false);

        var results = new List<ExcelWorkbookMergeResult>();
        foreach (var sourcePath in SourcePath)
        {
            if (string.IsNullOrWhiteSpace(sourcePath))
            {
                continue;
            }

            var resolvedSourcePath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(sourcePath);
            using var sourceDocument = ExcelDocumentService.LoadDocument(resolvedSourcePath, readOnly: true, autoSave: false);
            var result = targetDocument.MergeWorkbookFrom(sourceDocument, new ExcelWorkbookMergeOptions
            {
                SheetNames = SheetName,
                SheetNamePrefix = SheetNamePrefix,
                SheetNameValidationMode = ValidationMode,
                CopyMode = CopyMode
            });

            results.Add(result);
        }

        targetDocument.Save();
        foreach (var result in results)
        {
            WriteObject(result);
        }
    }
}
