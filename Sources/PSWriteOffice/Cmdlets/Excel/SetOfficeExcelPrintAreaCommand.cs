using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets the print area for a worksheet.</summary>
/// <example>
///   <summary>Set a sheet print area.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Set-OfficeExcelPrintArea -Path .\Report.xlsx -Sheet Data -Range A1:H100</code>
///   <para>Stores the worksheet-local Excel print area definition.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelPrintArea", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelPrintArea")]
public sealed class SetOfficeExcelPrintAreaCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("Path", "FilePath")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name. Defaults to the current sheet inside an ExcelSheet block.</summary>
    [Parameter]
    [Alias("WorksheetName")]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index when using a workbook object or path.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>A1 range to print.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Range { get; set; } = string.Empty;

    /// <summary>Emit the worksheet after setting the print area.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, InputPath, Document, readOnly: false);
        var document = workbook.Document;
        var sheet = ExcelWorkbookCommandService.ResolveSheet(this, document, ParameterSetName, Sheet, SheetIndex);
        document.SetPrintArea(sheet, Range, save: false);
        workbook.SaveIfOwned();

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }
}
