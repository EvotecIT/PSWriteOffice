using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets defined names (named ranges) from an Excel workbook.</summary>
/// <example>
///   <summary>List named ranges.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeExcelNamedRange -Path .\report.xlsx</code>
///   <para>Returns workbook-level named ranges.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelNamedRange", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeExcelNamedRangeCommand : PSCmdlet
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

    /// <summary>Optional named range to retrieve.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Optional sheet name for sheet-scoped names.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Optional sheet index (0-based) for sheet-scoped names.</summary>
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

            var scope = ResolveSheet(document);

            if (!string.IsNullOrWhiteSpace(Name))
            {
                var range = document.GetNamedRange(Name!, scope);
                if (range != null)
                {
                    WriteObject(CreateRecord(Name!, range, scope, ParameterSetName == ParameterSetPath ? InputPath : null));
                }
                return;
            }

            var ranges = document.GetAllNamedRanges(scope);
            foreach (var entry in ranges)
            {
                WriteObject(CreateRecord(entry.Key, entry.Value, scope, ParameterSetName == ParameterSetPath ? InputPath : null));
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

    private ExcelSheet? ResolveSheet(ExcelDocument document)
    {
        if (!string.IsNullOrWhiteSpace(Sheet))
        {
            return document[Sheet!];
        }

        if (SheetIndex.HasValue)
        {
            if (SheetIndex.Value < 0 || SheetIndex.Value >= document.Sheets.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(SheetIndex), "SheetIndex is out of range.");
            }
            return document.Sheets[SheetIndex.Value];
        }

        return null;
    }

    private static PSObject CreateRecord(string name, string range, ExcelSheet? scope, string? path)
    {
        var record = new PSObject();
        var sheetName = scope?.Name;
        record.Properties.Add(new PSNoteProperty("Name", name));
        record.Properties.Add(new PSNoteProperty("Range", NormalizeRange(range)));
        record.Properties.Add(new PSNoteProperty("Scope", sheetName ?? "Workbook"));
        if (!string.IsNullOrWhiteSpace(sheetName))
        {
            record.Properties.Add(new PSNoteProperty("Sheet", sheetName));
            record.Properties.Add(new PSNoteProperty("WorksheetName", sheetName));
        }
        if (!string.IsNullOrWhiteSpace(path))
        {
            record.Properties.Add(new PSNoteProperty("Path", path));
            record.Properties.Add(new PSNoteProperty("InputPath", path));
        }
        return record;
    }

    private static string NormalizeRange(string range)
    {
        return string.IsNullOrWhiteSpace(range)
            ? range
            : range.Replace("$", string.Empty);
    }
}
