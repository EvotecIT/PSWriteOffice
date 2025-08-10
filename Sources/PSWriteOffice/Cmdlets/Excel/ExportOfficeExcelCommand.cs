using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using ClosedXML.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsData.Export, "OfficeExcel")]
public class ExportOfficeExcelCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    [ValidateNotNullOrEmpty]
    public string FilePath { get; set; } = string.Empty;

    [Parameter]
    [Alias("Name")]
    public string WorksheetName { get; set; } = "Sheet1";

    [Parameter(ValueFromPipeline = true)]
    [Alias("TargetData")]
    public PSObject[] DataTable { get; set; } = Array.Empty<PSObject>();

    [Parameter]
    public int Row { get; set; } = 1;

    [Parameter]
    public int Column { get; set; } = 1;

    [Parameter]
    public SwitchParameter Append { get; set; }

    [Parameter]
    public SwitchParameter Show { get; set; }

    [Parameter]
    public SwitchParameter AllProperties { get; set; }

    [Parameter]
    public XLTransposeOptions? Transpose { get; set; }

    [Parameter]
    public SwitchParameter ShowRowStripes { get; set; }

    [Parameter]
    public SwitchParameter ShowColumnStripes { get; set; }

    [Parameter]
    public SwitchParameter DisableAutoFilter { get; set; }

    [Parameter]
    public SwitchParameter HideHeaderRow { get; set; }

    [Parameter]
    public SwitchParameter ShowTotalsRow { get; set; }

    [Parameter]
    public SwitchParameter EmphasizeFirstColumn { get; set; }

    [Parameter]
    public SwitchParameter EmphasizeLastColumn { get; set; }

    [Parameter]
    public SwitchParameter AutoSize { get; set; }

    [Parameter]
    public SwitchParameter FreezeTopRow { get; set; }

    [Parameter]
    public SwitchParameter FreezeFirstColumn { get; set; }

    [Parameter]
    public XLTableTheme Theme { get; set; } = XLTableTheme.None;

    [Parameter]
    public Hashtable? Formulas { get; set; }

    [Parameter]
    public PSObject[] PivotTables { get; set; } = Array.Empty<PSObject>();

    [Parameter]
    public PSObject[] Charts { get; set; } = Array.Empty<PSObject>();

    private readonly List<PSObject> _data = new();

    protected override void ProcessRecord()
    {
        if (DataTable != null)
        {
            foreach (var item in DataTable)
            {
                _data.Add(item);
            }
        }
    }

    protected override void EndProcessing()
    {
        try
        {
            var workbook = File.Exists(FilePath)
                ? ExcelDocumentService.LoadWorkbook(FilePath)
                : ExcelDocumentService.CreateWorkbook();

            var worksheet = ExcelDocumentService.AddWorksheet(
                workbook,
                WorksheetName,
                Append ? WorksheetExistOption.Skip : WorksheetExistOption.Replace);

            List<IDictionary<string, object?>> tableData;

            if (AllProperties)
            {
                var propertyNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var item in _data)
                {
                    if (item.BaseObject is IDictionary<string, object?> dict)
                    {
                        foreach (var key in dict.Keys)
                        {
                            propertyNames.Add(key);
                        }
                    }
                    else
                    {
                        foreach (var prop in item.Properties)
                        {
                            propertyNames.Add(prop.Name);
                        }
                    }
                }

                tableData = _data.Select(item =>
                {
                    IDictionary<string, object?> dict;
                    if (item.BaseObject is IDictionary<string, object?> existing)
                    {
                        dict = new Dictionary<string, object?>(existing, StringComparer.OrdinalIgnoreCase);
                    }
                    else
                    {
                        var psobj = item;
                        dict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                        foreach (var prop in psobj.Properties)
                        {
                            dict[prop.Name] = prop.Value;
                        }
                    }

                    foreach (var name in propertyNames)
                    {
                        if (!dict.ContainsKey(name))
                        {
                            dict[name] = null;
                        }
                    }

                    return dict;
                }).ToList();
            }
            else
            {
                tableData = _data.Select(item =>
                {
                    if (item.BaseObject is IDictionary<string, object?> dict)
                    {
                        return dict;
                    }

                    var psobj = item;
                    var result = new Dictionary<string, object?>();
                    foreach (var prop in psobj.Properties)
                    {
                        result[prop.Name] = prop.Value;
                    }
                    return (IDictionary<string, object?>)result;
                }).ToList();
            }

            if (Append && worksheet.Tables.Any())
            {
                // For append mode, we need to append to the existing table
                var existingTable = worksheet.Tables.First();
                var lastRow = existingTable.DataRange != null 
                    ? existingTable.DataRange.LastRow().RowNumber() + 1
                    : existingTable.RangeAddress.LastAddress.RowNumber + 1;
                
                // Add data rows below the existing table
                var rowIndex = 0;
                foreach (var row in tableData)
                {
                    var colIndex = existingTable.RangeAddress.FirstAddress.ColumnNumber;
                    foreach (var kvp in row)
                    {
                        worksheet.Cell(lastRow + rowIndex, colIndex).Value = XLCellValue.FromObject(kvp.Value);
                        colIndex++;
                    }
                    rowIndex++;
                }
                
                // Resize the table to include the new rows
                if (rowIndex > 0)
                {
                    var newLastRow = lastRow + rowIndex - 1;
                    existingTable.Resize(existingTable.RangeAddress.FirstAddress, 
                                        worksheet.Cell(newLastRow, existingTable.RangeAddress.LastAddress.ColumnNumber).Address);
                }
            }
            else if (Append && !worksheet.Tables.Any())
            {
                // If append mode but no table exists, create one
                ExcelDocumentService.InsertTable(
                    worksheet,
                    tableData,
                    Row,
                    Column,
                    Theme,
                    ShowRowStripes,
                    ShowColumnStripes,
                    !DisableAutoFilter,
                    !HideHeaderRow,
                    ShowTotalsRow,
                    EmphasizeFirstColumn,
                    EmphasizeLastColumn,
                    Transpose);
            }
            else
            {
                ExcelDocumentService.InsertTable(
                    worksheet,
                    tableData,
                    Row,
                    Column,
                    Theme,
                    ShowRowStripes,
                    ShowColumnStripes,
                    !DisableAutoFilter,
                    !HideHeaderRow,
                    ShowTotalsRow,
                    EmphasizeFirstColumn,
                    EmphasizeLastColumn,
                    Transpose);
            }

            if (Formulas != null && Formulas.Count > 0)
            {
                var formulas = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (DictionaryEntry entry in Formulas)
                {
                    if (entry.Key is string key && entry.Value != null)
                    {
                        formulas[key] = entry.Value.ToString()!;
                    }
                }
                ExcelDocumentService.ApplyFormulas(worksheet, formulas);
            }

            if (PivotTables != null && PivotTables.Length > 0)
            {
                foreach (var item in PivotTables)
                {
                    var dict = item.BaseObject as IDictionary<string, object?> ?? item.Properties.ToDictionary(p => p.Name, p => p.Value, StringComparer.OrdinalIgnoreCase);
                    var name = dict.TryGetValue("Name", out var n) ? n?.ToString() ?? "PivotTable1" : "PivotTable1";
                    var sourceRange = dict.TryGetValue("SourceRange", out var sr) ? sr?.ToString() ?? string.Empty : string.Empty;
                    var targetCell = dict.TryGetValue("TargetCell", out var tc) ? tc?.ToString() ?? "A1" : "A1";

                    IEnumerable<string>? rowFields = null;
                    if (dict.TryGetValue("RowFields", out var rf) && rf is IEnumerable<object?> rfEnum)
                    {
                        rowFields = rfEnum.Select(o => o?.ToString() ?? string.Empty).Where(s => !string.IsNullOrEmpty(s)).ToList();
                    }

                    IEnumerable<string>? columnFields = null;
                    if (dict.TryGetValue("ColumnFields", out var cf) && cf is IEnumerable<object?> cfEnum)
                    {
                        columnFields = cfEnum.Select(o => o?.ToString() ?? string.Empty).Where(s => !string.IsNullOrEmpty(s)).ToList();
                    }

                    IDictionary<string, XLPivotSummary>? values = null;
                    if (dict.TryGetValue("Values", out var v) && v is IDictionary vDict)
                    {
                        values = new Dictionary<string, XLPivotSummary>(StringComparer.OrdinalIgnoreCase);
                        foreach (DictionaryEntry entry in vDict)
                        {
                            var summary = Enum.TryParse<XLPivotSummary>(entry.Value?.ToString(), true, out var res)
                                ? res
                                : XLPivotSummary.Sum;
                            values[entry.Key.ToString()!] = summary;
                        }
                    }

                    ExcelDocumentService.AddPivotTable(worksheet, name, sourceRange, targetCell, rowFields, columnFields, values);
                }
            }

            if (AutoSize)
            {
                ExcelDocumentService.AutoSizeColumns(worksheet);
            }

            if (FreezeTopRow)
            {
                ExcelDocumentService.FreezeTopRow(worksheet);
            }

            if (FreezeFirstColumn)
            {
                ExcelDocumentService.FreezeFirstColumn(worksheet);
            }

            ExcelDocumentService.SaveWorkbook(workbook, FilePath, Show);

            if (Charts != null && Charts.Length > 0)
            {
                foreach (var chart in Charts)
                {
                    var dict = chart.BaseObject as IDictionary<string, object?> ?? chart.Properties.ToDictionary(p => p.Name, p => p.Value, StringComparer.OrdinalIgnoreCase);
                    var title = dict.TryGetValue("Title", out var t) ? t?.ToString() ?? string.Empty : string.Empty;
                    ExcelDocumentService.AddChart(FilePath, WorksheetName, title);
                }
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ExportOfficeExcelFailed", ErrorCategory.NotSpecified, FilePath));
        }
    }
}
