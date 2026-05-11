using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets a compact structural summary of an Excel workbook.</summary>
/// <example>
///   <summary>Summarize workbook contents.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeExcelSummary -Path .\report.xlsx</code>
///   <para>Returns workbook-level counts plus per-sheet tables, charts, pivots, links, comments, and used ranges.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelSummary", DefaultParameterSetName = ParameterSetPath)]
[Alias("ExcelSummary")]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeExcelSummaryCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the workbook.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath", "Path")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Workbook to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Include per-sheet details in the returned object.</summary>
    [Parameter]
    public SwitchParameter IncludeSheets { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDocument? loadedDocument = null;
        var dispose = false;

        try
        {
            SpreadsheetDocument spreadsheet;
            string? path;

            if (ParameterSetName == ParameterSetPath)
            {
                var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
                if (!File.Exists(resolvedPath))
                {
                    throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
                }

                loadedDocument = ExcelDocumentService.LoadDocument(resolvedPath, readOnly: true, autoSave: false);
                spreadsheet = loadedDocument._spreadSheetDocument;
                path = resolvedPath;
                dispose = true;
            }
            else
            {
                if (Document == null)
                {
                    throw new InvalidOperationException("Excel workbook was not provided.");
                }

                spreadsheet = Document._spreadSheetDocument;
                path = Document.FilePath;
            }

            WriteObject(CreateSummary(spreadsheet, path, IncludeSheets.IsPresent));
        }
        finally
        {
            if (dispose)
            {
                loadedDocument?.Dispose();
            }
        }
    }

    private static PSObject CreateSummary(SpreadsheetDocument spreadsheet, string? path, bool includeSheets)
    {
        var workbookPart = spreadsheet.WorkbookPart ?? throw new InvalidOperationException("Workbook part was not found.");
        var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook was not found.");
        var sheets = workbook.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
        var sheetSummaries = sheets.Select((sheet, index) => CreateSheetSummary(workbookPart, sheet, index + 1)).ToArray();
        var namedRangeCount = workbook.DefinedNames?.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>().Count() ?? 0;

        var summary = new PSObject();
        summary.Properties.Add(new PSNoteProperty("Path", path));
        summary.Properties.Add(new PSNoteProperty("SheetCount", sheetSummaries.Length));
        summary.Properties.Add(new PSNoteProperty("VisibleSheetCount", sheetSummaries.Count(IsVisibleSheet)));
        summary.Properties.Add(new PSNoteProperty("HiddenSheetCount", sheetSummaries.Count(IsHiddenSheet)));
        summary.Properties.Add(new PSNoteProperty("VeryHiddenSheetCount", sheetSummaries.Count(IsVeryHiddenSheet)));
        summary.Properties.Add(new PSNoteProperty("TableCount", sheetSummaries.Sum(GetIntProperty("TableCount"))));
        summary.Properties.Add(new PSNoteProperty("ChartCount", sheetSummaries.Sum(GetIntProperty("ChartCount"))));
        summary.Properties.Add(new PSNoteProperty("PivotTableCount", sheetSummaries.Sum(GetIntProperty("PivotTableCount"))));
        summary.Properties.Add(new PSNoteProperty("SparklineGroupCount", sheetSummaries.Sum(GetIntProperty("SparklineGroupCount"))));
        summary.Properties.Add(new PSNoteProperty("HyperlinkCount", sheetSummaries.Sum(GetIntProperty("HyperlinkCount"))));
        summary.Properties.Add(new PSNoteProperty("CommentCount", sheetSummaries.Sum(GetIntProperty("CommentCount"))));
        summary.Properties.Add(new PSNoteProperty("NamedRangeCount", namedRangeCount));

        if (includeSheets)
        {
            summary.Properties.Add(new PSNoteProperty("Sheets", sheetSummaries));
        }

        return summary;
    }

    private static PSObject CreateSheetSummary(WorkbookPart workbookPart, Sheet sheet, int index)
    {
        var state = NormalizeSheetState(sheet.State?.InnerText);
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("Index", index));
        record.Properties.Add(new PSNoteProperty("Name", sheet.Name?.Value ?? string.Empty));
        record.Properties.Add(new PSNoteProperty("State", state));

        if (sheet.Id?.Value == null)
        {
            AddEmptySheetCounts(record);
            return record;
        }

        var sheetPart = workbookPart.GetPartById(sheet.Id.Value);
        if (sheetPart is ChartsheetPart chartsheetPart)
        {
            AddChartSheetCounts(record, chartsheetPart);
            return record;
        }

        if (sheetPart is not WorksheetPart worksheetPart)
        {
            AddEmptySheetCounts(record);
            return record;
        }

        var worksheet = worksheetPart.Worksheet ?? new Worksheet();
        var tableRecords = GetTableRecords(worksheetPart).ToArray();
        var chartCount = CountChartParts(worksheetPart);
        var pivotCount = worksheetPart.PivotTableParts.Count();
        var sparklineGroupCount = worksheet.Descendants<SparklineGroups>().Sum(groups => groups.Elements<SparklineGroup>().Count());
        var hyperlinkCount = worksheet.Elements<Hyperlinks>().FirstOrDefault()?.Elements<Hyperlink>().Count() ?? 0;
        var commentCount = worksheetPart.WorksheetCommentsPart?.Comments?.CommentList?.Elements<Comment>().Count() ?? 0;

        record.Properties.Add(new PSNoteProperty("UsedRange", worksheet.SheetDimension?.Reference?.Value));
        record.Properties.Add(new PSNoteProperty("TableCount", tableRecords.Length));
        record.Properties.Add(new PSNoteProperty("ChartCount", chartCount));
        record.Properties.Add(new PSNoteProperty("PivotTableCount", pivotCount));
        record.Properties.Add(new PSNoteProperty("SparklineGroupCount", sparklineGroupCount));
        record.Properties.Add(new PSNoteProperty("HyperlinkCount", hyperlinkCount));
        record.Properties.Add(new PSNoteProperty("CommentCount", commentCount));
        record.Properties.Add(new PSNoteProperty("Tables", tableRecords));
        return record;
    }

    private static void AddChartSheetCounts(PSObject record, ChartsheetPart chartsheetPart)
    {
        record.Properties.Add(new PSNoteProperty("UsedRange", null));
        record.Properties.Add(new PSNoteProperty("TableCount", 0));
        record.Properties.Add(new PSNoteProperty("ChartCount", CountChartParts(chartsheetPart)));
        record.Properties.Add(new PSNoteProperty("PivotTableCount", 0));
        record.Properties.Add(new PSNoteProperty("SparklineGroupCount", 0));
        record.Properties.Add(new PSNoteProperty("HyperlinkCount", 0));
        record.Properties.Add(new PSNoteProperty("CommentCount", 0));
        record.Properties.Add(new PSNoteProperty("Tables", Array.Empty<PSObject>()));
    }

    private static int CountChartParts(OpenXmlPartContainer container)
    {
        return container.Parts.Sum(part =>
            (part.OpenXmlPart is ChartPart ? 1 : 0) + CountChartParts(part.OpenXmlPart));
    }

    private static IEnumerable<PSObject> GetTableRecords(WorksheetPart worksheetPart)
    {
        foreach (var tableDefinitionPart in worksheetPart.TableDefinitionParts)
        {
            var table = tableDefinitionPart.Table;
            if (table == null)
            {
                continue;
            }

            var record = new PSObject();
            record.Properties.Add(new PSNoteProperty("Name", table.Name?.Value ?? table.DisplayName?.Value));
            record.Properties.Add(new PSNoteProperty("Range", table.Reference?.Value));
            record.Properties.Add(new PSNoteProperty("Style", table.TableStyleInfo?.Name?.Value));
            yield return record;
        }
    }

    private static void AddEmptySheetCounts(PSObject record)
    {
        record.Properties.Add(new PSNoteProperty("UsedRange", null));
        record.Properties.Add(new PSNoteProperty("TableCount", 0));
        record.Properties.Add(new PSNoteProperty("ChartCount", 0));
        record.Properties.Add(new PSNoteProperty("PivotTableCount", 0));
        record.Properties.Add(new PSNoteProperty("SparklineGroupCount", 0));
        record.Properties.Add(new PSNoteProperty("HyperlinkCount", 0));
        record.Properties.Add(new PSNoteProperty("CommentCount", 0));
        record.Properties.Add(new PSNoteProperty("Tables", Array.Empty<PSObject>()));
    }

    private static Func<PSObject, int> GetIntProperty(string name)
    {
        return record => record.Properties[name]?.Value is int value ? value : 0;
    }

    private static string NormalizeSheetState(string? state)
    {
        if (string.IsNullOrWhiteSpace(state))
        {
            return "Visible";
        }

        return state!.Equals("veryHidden", StringComparison.OrdinalIgnoreCase)
            ? "VeryHidden"
            : state.Equals("hidden", StringComparison.OrdinalIgnoreCase)
                ? "Hidden"
                : "Visible";
    }

    private static bool IsVisibleSheet(PSObject record)
    {
        return string.Equals(record.Properties["State"]?.Value as string, "Visible", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsHiddenSheet(PSObject record)
    {
        return string.Equals(record.Properties["State"]?.Value as string, "Hidden", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsVeryHiddenSheet(PSObject record)
    {
        return string.Equals(record.Properties["State"]?.Value as string, "VeryHidden", StringComparison.OrdinalIgnoreCase);
    }
}
