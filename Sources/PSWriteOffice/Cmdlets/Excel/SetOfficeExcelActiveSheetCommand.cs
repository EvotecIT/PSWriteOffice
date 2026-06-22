using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets the worksheet that opens as the active sheet.</summary>
/// <example>
///   <summary>Open the workbook on the Summary sheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$sheet = Set-OfficeExcelActiveSheet -Path .\Report.xlsx -Sheet Summary -PassThru
/// Get-OfficeExcelWorksheetView -Path .\Report.xlsx -Sheet $sheet.Name |
///     Select-Object SheetName, View, TopLeftCell</code>
///   <para>Updates workbook view state so spreadsheet applications open on Summary.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelActiveSheet", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelActiveSheet")]
public sealed class SetOfficeExcelActiveSheetCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("Path", "FilePath")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name to activate. Defaults to the current sheet inside an ExcelSheet block.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Parameter(ParameterSetName = ParameterSetPath)]
    [Alias("WorksheetName")]
    public string? Sheet { get; set; }

    /// <summary>Zero-based worksheet index to activate.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Parameter(ParameterSetName = ParameterSetPath)]
    public int? SheetIndex { get; set; }

    /// <summary>Emit the activated worksheet.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (!string.IsNullOrWhiteSpace(Sheet) && SheetIndex.HasValue)
        {
            throw new PSArgumentException("Specify either -Sheet or -SheetIndex, not both.");
        }

        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, InputPath, Document, readOnly: false);
        ExcelSheet sheet = ResolveTargetSheet(workbook.Document);
        workbook.Document.SetActiveWorksheet(sheet);
        workbook.SaveIfOwned();

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private ExcelSheet ResolveTargetSheet(ExcelDocument document)
    {
        if (ParameterSetName == ParameterSetContext)
        {
            var context = ExcelDslContext.Require(this);
            return context.RequireSheet();
        }

        if (!string.IsNullOrWhiteSpace(Sheet))
        {
            return ExcelSheetResolver.Resolve(document, Sheet, null);
        }

        if (SheetIndex.HasValue)
        {
            return ExcelSheetResolver.Resolve(document, null, SheetIndex);
        }

        throw new PSArgumentException("Specify -Sheet or -SheetIndex.");
    }
}
