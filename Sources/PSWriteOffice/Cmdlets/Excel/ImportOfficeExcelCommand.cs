using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using PSWriteOffice.Services.Excel;
using ValidateScriptAttribute = PSWriteOffice.Validation.ValidateScriptAttribute;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsData.Import, "OfficeExcel")]
public class ImportOfficeExcelCommand : PSCmdlet
{
    [Alias("LiteralPath")]
    [Parameter(Mandatory = true)]
    [ValidateNotNullOrEmpty]
    [ValidateScript("{ Test-Path $_ }")]
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
                WriteObject(single, !AsDataTable.IsPresent);
            }
            else if ((WorkSheetName == null || WorkSheetName.Length == 0) && converted.Count == 1)
            {
                WriteObject(converted.Values.First(), !AsDataTable.IsPresent);
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
