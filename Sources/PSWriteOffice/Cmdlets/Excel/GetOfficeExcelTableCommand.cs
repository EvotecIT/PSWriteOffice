using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets Excel tables defined in a workbook.</summary>
/// <example>
///   <summary>List tables in a workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeExcelTable -Path .\report.xlsx</code>
///   <para>Returns table metadata (name, range, sheet).</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelTable", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeExcelTableCommand : PSCmdlet
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

    /// <summary>Optional table name filter.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Optional sheet name filter.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Optional sheet index (0-based) filter.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDocument? document = null;
        var dispose = false;

        try
        {
            if (ParameterSetName == ParameterSetPath)
            {
                var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
                if (!File.Exists(resolvedPath))
                {
                    throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
                }
                document = ExcelDocumentService.LoadDocument(resolvedPath, readOnly: true, autoSave: false);
                dispose = true;
            }
            else
            {
                document = Document;
            }

            if (document == null)
            {
                throw new InvalidOperationException("Excel workbook was not provided.");
            }

            var sheetFilter = ResolveSheetName(document);
            var spreadsheet = document._spreadSheetDocument;
            if (spreadsheet?.WorkbookPart == null)
            {
                return;
            }

            var workbookPart = spreadsheet.WorkbookPart;
            var sheetLookup = BuildSheetLookup(workbookPart);

            foreach (var worksheetPart in workbookPart.WorksheetParts)
            {
                var relId = workbookPart.GetIdOfPart(worksheetPart);
                if (string.IsNullOrWhiteSpace(relId))
                {
                    continue;
                }
                sheetLookup.TryGetValue(relId, out var sheetName);
                sheetName ??= string.Empty;

                if (!string.IsNullOrWhiteSpace(sheetFilter) &&
                    !string.Equals(sheetName, sheetFilter, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                foreach (var tablePart in worksheetPart.TableDefinitionParts)
                {
                    var table = tablePart.Table;
                    if (table == null)
                    {
                        continue;
                    }

                    var tableName = table.Name?.Value ?? table.DisplayName?.Value ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(Name) &&
                        !string.Equals(tableName, Name, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var range = table.Reference?.Value ?? string.Empty;
                    WriteObject(CreateRecord(tableName, range, sheetName));
                }
            }
        }
        finally
        {
            if (dispose)
            {
                document?.Dispose();
            }
        }
    }

    private string? ResolveSheetName(ExcelDocument document)
    {
        if (!string.IsNullOrWhiteSpace(Sheet))
        {
            return Sheet;
        }

        if (SheetIndex.HasValue)
        {
            if (SheetIndex.Value < 0 || SheetIndex.Value >= document.Sheets.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(SheetIndex), "SheetIndex is out of range.");
            }
            return document.Sheets[SheetIndex.Value].Name;
        }

        return null;
    }

    private static Dictionary<string, string> BuildSheetLookup(WorkbookPart workbookPart)
    {
        var lookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var sheets = workbookPart.Workbook.Sheets?.OfType<Sheet>() ?? Enumerable.Empty<Sheet>();
        foreach (var sheet in sheets)
        {
            var id = sheet.Id?.Value;
            var name = sheet.Name?.Value;
            if (id == null || name == null)
            {
                continue;
            }
            if (id.Length == 0 || name.Length == 0)
            {
                continue;
            }
            lookup[id] = name;
        }
        return lookup;
    }

    private static PSObject CreateRecord(string name, string range, string sheet)
    {
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("Name", name));
        record.Properties.Add(new PSNoteProperty("Range", range));
        record.Properties.Add(new PSNoteProperty("Sheet", sheet));
        return record;
    }
}
