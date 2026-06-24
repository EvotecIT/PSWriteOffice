using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Imports selected or all worksheets from one Excel workbook into another.</summary>
/// <example>
///   <summary>Import all source sheets with a prefix.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$merge = Join-OfficeExcelWorkbook -Path .\Target.xlsx -SourcePath .\Source.xlsx -SourceSheet Data,Summary -SheetNamePrefix 'Imported '
/// Get-OfficeExcelSummary -Path .\Target.xlsx |
///     Select-Object Path, WorksheetCount</code>
///   <para>Copies worksheets from Source.xlsx into Target.xlsx using OfficeIMO workbook merge logic.</para>
/// </example>
[Cmdlet(VerbsCommon.Join, "OfficeExcelWorkbook", DefaultParameterSetName = ParameterSetPath)]
[Alias("Merge-OfficeExcelWorkbook", "ExcelWorkbookJoin", "ExcelWorkbookMerge")]
[OutputType(typeof(ExcelWorkbookMergeResult))]
public sealed class JoinOfficeExcelWorkbookCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Target workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("Path", "FilePath")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Target workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Optional source workbook object.</summary>
    [Parameter]
    public ExcelDocument? SourceDocument { get; set; }

    /// <summary>Optional source workbook path.</summary>
    [Parameter]
    public string? SourcePath { get; set; }

    /// <summary>Specific source worksheet names to import. Defaults to all source sheets.</summary>
    [Parameter]
    public string[]? SourceSheet { get; set; }

    /// <summary>Prefix added to every imported worksheet name.</summary>
    [Parameter]
    public string? SheetNamePrefix { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (SourceDocument == null && string.IsNullOrWhiteSpace(SourcePath))
        {
            throw new PSArgumentException("Provide SourceDocument or SourcePath.");
        }

        using var targetWorkbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, InputPath, Document, readOnly: false);
        using var sourceWorkbook = ExcelWorkbookCommandService.ResolveSourceWorkbook(this, targetWorkbook.Document, SourceDocument, SourcePath, readOnly: true);

        ExcelWorkbookMergeResult result = targetWorkbook.Document.JoinWorkbookFrom(sourceWorkbook.Document, new ExcelWorkbookMergeOptions
        {
            SheetNames = SourceSheet,
            SheetNamePrefix = SheetNamePrefix
        });

        targetWorkbook.SaveIfOwned();
        WriteObject(result);
    }
}
