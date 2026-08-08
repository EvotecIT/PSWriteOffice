using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Exports one named worksheet chart as an image file.</summary>
/// <example>
///   <summary>Export a chart for inline use in an email.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Export-OfficeExcelChartImage -Path .\Report.xlsx -WorksheetName Summary -ChartName TicketStatus -OutputPath .\ticket-status.png</code>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeExcelChartImage", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeImageExportResult))]
public sealed class ExportOfficeExcelChartImageCommand : PSCmdlet
{
    /// <summary>Path to the workbook.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open workbook instance.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document")]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Name of the worksheet containing the chart.</summary>
    [Parameter(Mandatory = true)]
    public string WorksheetName { get; set; } = string.Empty;

    /// <summary>Chart name as stored in the worksheet drawing layer.</summary>
    [Parameter(Mandatory = true)]
    public string ChartName { get; set; } = string.Empty;

    /// <summary>Optional destination image file. When omitted, returns the in-memory image result only.</summary>
    [Parameter(Position = 1)]
    public string? OutputPath { get; set; }

    /// <summary>Output image format.</summary>
    [Parameter]
    public OfficeImageExportFormat Format { get; set; } = OfficeImageExportFormat.Png;

    /// <summary>Optional rendering, size, font, and diagnostic policy settings.</summary>
    [Parameter]
    public ExcelImageExportOptions? Options { get; set; }

    /// <summary>Replace an existing destination file.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDocument? owned = null;
        try
        {
            ExcelDocument document = Document;
            if (ParameterSetName == "Path")
            {
                string input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
                owned = ExcelDocumentService.LoadDocument(input, readOnly: true, autoSave: false);
                document = owned;
            }

            ExcelChart chart = document[WorksheetName].GetChart(ChartName)
                ?? throw new ItemNotFoundException(
                    $"Chart '{ChartName}' was not found on worksheet '{WorksheetName}'.");
            OfficeImageExportResult result = chart
                .ExportImage(Format, Options);

            if (!string.IsNullOrWhiteSpace(OutputPath))
            {
                string output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
                if (!ShouldProcess(output, $"Save Excel chart {WorksheetName}!{ChartName} as {Format}")) return;
                result.Save(output, Force.IsPresent
                    ? OfficeImageExportFileConflictPolicy.Replace
                    : OfficeImageExportFileConflictPolicy.FailIfExists);
            }
            WriteObject(result);
        }
        finally
        {
            owned?.Dispose();
        }
    }
}
