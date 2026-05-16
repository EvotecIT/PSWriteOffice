using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Text.RegularExpressions;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Finds text in worksheet values.</summary>
/// <example>
///   <summary>Find values in a workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Find-OfficeExcel -Path .\Report.xlsx -Text Ready</code>
///   <para>Returns matching cells with sheet, address, row, column, and value metadata.</para>
/// </example>
[Cmdlet(VerbsCommon.Find, "OfficeExcel", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(PSObject))]
public sealed class FindOfficeExcelCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to inspect.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("Path", "FilePath")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Workbook to inspect outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Text or pattern to find.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Worksheet name. Defaults to all sheets for path/document use and current sheet inside an ExcelSheet block.</summary>
    [Parameter]
    [Alias("WorksheetName")]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index when using a workbook object or path.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>A1 range to search. Defaults to each selected worksheet's used range.</summary>
    [Parameter]
    public string? Range { get; set; }

    /// <summary>Use case-sensitive matching.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <summary>Treat -Text as a regular expression.</summary>
    [Parameter]
    public SwitchParameter Regex { get; set; }

    /// <summary>Require an exact cell text match instead of substring matching.</summary>
    [Parameter]
    public SwitchParameter Exact { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDocument? document = null;
        var dispose = false;

        try
        {
            document = ResolveDocument(out dispose);
            foreach (var sheet in ResolveSheets(document))
            {
                var range = string.IsNullOrWhiteSpace(Range) ? sheet.GetUsedRangeA1() : Range!;
                using var reader = document.CreateReader();
                var sheetReader = reader.GetSheet(sheet.Name);
                foreach (var cell in sheetReader.EnumerateRange(range))
                {
                    var cellText = Convert.ToString(cell.Value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
                    if (IsMatch(cellText))
                    {
                        WriteObject(CreateRecord(sheet.Name, cell.Row, cell.Column, cell.Value));
                    }
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

    private ExcelDocument ResolveDocument(out bool dispose)
    {
        dispose = false;
        if (ParameterSetName == ParameterSetPath)
        {
            var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
            if (!File.Exists(resolvedPath))
            {
                throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
            }

            dispose = true;
            return ExcelDocumentService.LoadDocument(resolvedPath, readOnly: true, autoSave: false);
        }

        return ParameterSetName == ParameterSetDocument
            ? Document ?? throw new PSArgumentException("Provide an Excel document.")
            : ExcelDslContext.Require(this).Document;
    }

    private IEnumerable<ExcelSheet> ResolveSheets(ExcelDocument document)
    {
        if (ParameterSetName == ParameterSetContext && string.IsNullOrWhiteSpace(Sheet) && !SheetIndex.HasValue)
        {
            yield return ExcelDslContext.Require(this).RequireSheet();
            yield break;
        }

        if (!string.IsNullOrWhiteSpace(Sheet) || SheetIndex.HasValue)
        {
            yield return ExcelSheetResolver.Resolve(document, Sheet, SheetIndex);
            yield break;
        }

        foreach (var sheet in document.Sheets)
        {
            yield return sheet;
        }
    }

    private bool IsMatch(string value)
    {
        var comparison = CaseSensitive.IsPresent ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        if (Regex.IsPresent)
        {
            var options = CaseSensitive.IsPresent ? RegexOptions.None : RegexOptions.IgnoreCase;
            return System.Text.RegularExpressions.Regex.IsMatch(value, Text, options);
        }

        return Exact.IsPresent
            ? string.Equals(value, Text, comparison)
            : value.IndexOf(Text, comparison) >= 0;
    }

    private static PSObject CreateRecord(string sheetName, int row, int column, object? value)
    {
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("Sheet", sheetName));
        record.Properties.Add(new PSNoteProperty("Address", A1.CellReference(row, column)));
        record.Properties.Add(new PSNoteProperty("Row", row));
        record.Properties.Add(new PSNoteProperty("Column", column));
        record.Properties.Add(new PSNoteProperty("Value", value));
        return record;
    }
}
