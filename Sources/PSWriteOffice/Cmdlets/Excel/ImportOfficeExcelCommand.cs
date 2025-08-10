using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsData.Import, "OfficeExcel")]
public class ImportOfficeExcelCommand : PSCmdlet
{
    [Alias("LiteralPath")]
    [Parameter(Mandatory = true)]
    [ValidateNotNullOrEmpty]
    public string FilePath { get; set; } = string.Empty;

    [Parameter]
    public string[]? WorkSheetName { get; set; }

    [Parameter]
    public CultureInfo? Culture { get; set; }

    [Parameter]
    public int? StartRow { get; set; }

    [Parameter]
    public int? EndRow { get; set; }

    [Parameter]
    public int? StartColumn { get; set; }

    [Parameter]
    public int? EndColumn { get; set; }

    [Parameter]
    public int? HeaderRow { get; set; }

    [Parameter]
    public SwitchParameter NoHeader { get; set; }

    [Parameter]
    public Type? Type { get; set; }

    [Parameter]
    public SwitchParameter AsDataTable { get; set; }

    protected override void ProcessRecord()
    {
        // Validate file exists
        if (!File.Exists(FilePath))
        {
            var ex = new FileNotFoundException($"File not found: {FilePath}", FilePath);
            WriteError(new ErrorRecord(ex, "FileNotFound", ErrorCategory.ObjectNotFound, FilePath));
            return;
        }

        try
        {
            var raw = ExcelDocumentService.ImportWorkbook(FilePath, WorkSheetName, Culture, StartRow, EndRow, StartColumn, EndColumn, HeaderRow, NoHeader);
            var converted = new Dictionary<string, object>();
            foreach (var kvp in raw)
            {
                converted[kvp.Key] = ExcelDocumentService.ConvertWorksheetData(kvp.Value, Type, AsDataTable);
            }

            if (WorkSheetName != null && WorkSheetName.Length == 1 && converted.TryGetValue(WorkSheetName[0], out var single))
            {
                if (AsDataTable.IsPresent)
                {
                    WriteObject(single);
                }
                else
                {
                    // Convert dictionaries to PSObjects
                    if (single is IEnumerable<IDictionary<string, object?>> rows)
                    {
                        foreach (var row in rows)
                        {
                            var psObj = new PSObject();
                            foreach (var kvp in row)
                            {
                                psObj.Properties.Add(new PSNoteProperty(kvp.Key, kvp.Value));
                            }
                            WriteObject(psObj);
                        }
                    }
                    else
                    {
                        WriteObject(single, true);
                    }
                }
            }
            else if ((WorkSheetName == null || WorkSheetName.Length == 0) && converted.Count == 1)
            {
                var singleSheet = converted.Values.First();
                if (AsDataTable.IsPresent)
                {
                    WriteObject(singleSheet);
                }
                else
                {
                    // Convert dictionaries to PSObjects
                    if (singleSheet is IEnumerable<IDictionary<string, object?>> rows)
                    {
                        foreach (var row in rows)
                        {
                            var psObj = new PSObject();
                            foreach (var kvp in row)
                            {
                                psObj.Properties.Add(new PSNoteProperty(kvp.Key, kvp.Value));
                            }
                            WriteObject(psObj);
                        }
                    }
                    else
                    {
                        WriteObject(singleSheet, true);
                    }
                }
            }
            else
            {
                WriteObject(converted);
            }
        }
        catch (FileNotFoundException ex)
        {
            WriteError(new ErrorRecord(ex, "FileNotFound", ErrorCategory.ObjectNotFound, FilePath));
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ExcelImportFailed", ErrorCategory.InvalidOperation, FilePath));
        }
    }
}
