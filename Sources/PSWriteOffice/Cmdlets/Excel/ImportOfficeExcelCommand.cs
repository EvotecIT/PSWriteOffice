using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Imports rows from an Excel workbook as PowerShell objects.</summary>
/// <para>Provides an ImportExcel-style read command over the OfficeIMO reader pipeline.</para>
/// <example>
///   <summary>Import worksheet rows.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Import-OfficeExcel -Path .\Report.xlsx -WorksheetName Data</code>
///   <para>Reads the used range on the Data worksheet and emits PSCustomObjects.</para>
/// </example>
[Cmdlet(VerbsData.Import, "OfficeExcel")]
[Alias("ExcelImport")]
public sealed class ImportOfficeExcelCommand : PSCmdlet
{
    /// <summary>Workbook path to import.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Worksheet name to read; defaults to the first sheet.</summary>
    [Parameter]
    [Alias("Sheet")]
    public string? WorksheetName { get; set; }

    /// <summary>Zero-based worksheet index to read.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>Optional A1 range to read. When omitted, the used range is imported.</summary>
    [Parameter]
    public string? Range { get; set; }

    /// <summary>Starting row for an explicit rectangular range.</summary>
    [Parameter]
    public int? StartRow { get; set; }

    /// <summary>Ending row for an explicit rectangular range.</summary>
    [Parameter]
    public int? EndRow { get; set; }

    /// <summary>Starting column for an explicit rectangular range.</summary>
    [Parameter]
    public int? StartColumn { get; set; }

    /// <summary>Ending column for an explicit rectangular range.</summary>
    [Parameter]
    public int? EndColumn { get; set; }

    /// <summary>Treat all rows as data and generate column names instead of using the first row as headers.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Prefer decimals instead of doubles for numeric values.</summary>
    [Parameter]
    public SwitchParameter NumericAsDecimal { get; set; }

    /// <summary>Emit rows as hashtables instead of PSCustomObjects.</summary>
    [Parameter]
    public SwitchParameter AsHashtable { get; set; }

    /// <summary>Emit a DataTable instead of enumerating row objects.</summary>
    [Parameter]
    public SwitchParameter AsDataTable { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
        }

        if (!string.IsNullOrWhiteSpace(Range) && HasCoordinateRange())
        {
            throw new PSArgumentException("Specify either -Range or coordinate bounds, but not both.");
        }

        var options = ExcelReadOutputService.CreateOptions(NumericAsDecimal.IsPresent);
        using var reader = ExcelDocumentReader.Open(resolvedPath, options);
        var sheet = ExcelReadOutputService.ResolveSheetReader(reader, WorksheetName, SheetIndex);
        var range = ResolveRange(sheet);
        var table = sheet.ReadRangeAsDataTable(range, headersInFirstRow: !NoHeader.IsPresent);

        ExcelReadOutputService.WriteOutput(this, table, AsDataTable.IsPresent, AsHashtable.IsPresent);
    }

    private string ResolveRange(ExcelSheetReader sheet)
    {
        if (!string.IsNullOrWhiteSpace(Range))
        {
            return Range!;
        }

        if (HasCoordinateRange())
        {
            if (!StartRow.HasValue || !EndRow.HasValue || !StartColumn.HasValue || !EndColumn.HasValue)
            {
                throw new PSArgumentException("StartRow, EndRow, StartColumn, and EndColumn must all be provided when using coordinate bounds.");
            }

            if (StartRow.Value < 1 || EndRow.Value < 1 || StartColumn.Value < 1 || EndColumn.Value < 1)
            {
                throw new PSArgumentException("Coordinate bounds must be 1 or greater.");
            }

            if (StartRow.Value > EndRow.Value)
            {
                throw new PSArgumentException("StartRow must be less than or equal to EndRow.");
            }

            if (StartColumn.Value > EndColumn.Value)
            {
                throw new PSArgumentException("StartColumn must be less than or equal to EndColumn.");
            }

            return $"{A1.CellReference(StartRow.Value, StartColumn.Value)}:{A1.CellReference(EndRow.Value, EndColumn.Value)}";
        }

        return sheet.GetUsedRangeA1();
    }

    private bool HasCoordinateRange()
    {
        return StartRow.HasValue || EndRow.HasValue || StartColumn.HasValue || EndColumn.HasValue;
    }
}
