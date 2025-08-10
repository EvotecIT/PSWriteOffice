using System;
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
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ExportOfficeExcelFailed", ErrorCategory.NotSpecified, FilePath));
        }
    }
}
