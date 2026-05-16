using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Text.RegularExpressions;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Replaces text in worksheet values.</summary>
/// <example>
///   <summary>Replace status text in a workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Update-OfficeExcelText -Path .\Report.xlsx -OldValue Draft -NewValue Ready</code>
///   <para>Updates matching text cells and returns the replacement count.</para>
/// </example>
[Cmdlet(VerbsData.Update, "OfficeExcelText", DefaultParameterSetName = ParameterSetPath)]
[Alias("Replace-OfficeExcelText")]
[OutputType(typeof(int))]
public sealed class UpdateOfficeExcelTextCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("Path", "FilePath")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Text or pattern to replace.</summary>
    [Parameter(Mandatory = true)]
    public string OldValue { get; set; } = string.Empty;

    /// <summary>Replacement text.</summary>
    [Parameter(Mandatory = true)]
    public string NewValue { get; set; } = string.Empty;

    /// <summary>Worksheet name. Defaults to all sheets for path/document use and current sheet inside an ExcelSheet block.</summary>
    [Parameter]
    [Alias("WorksheetName")]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index when using a workbook object or path.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>A1 range to update. Defaults to each selected worksheet's used range.</summary>
    [Parameter]
    public string? Range { get; set; }

    /// <summary>Use case-sensitive matching.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <summary>Treat -OldValue as a regular expression.</summary>
    [Parameter]
    public SwitchParameter Regex { get; set; }

    /// <summary>Open the file after saving when using -Path.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    public SwitchParameter Show { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDocument? document = null;
        var dispose = false;
        var replacements = 0;

        try
        {
            document = ResolveDocument(out dispose);
            foreach (var sheet in ResolveSheets(document))
            {
                replacements += ReplaceInSheet(document, sheet);
            }

            if (dispose)
            {
                document.Save(false);
                var savedPath = document.FilePath ?? InputPath;
                document.Dispose();
                document = null;
                if (Show.IsPresent)
                {
                    FileOpenService.Open(savedPath);
                }
            }

            WriteObject(replacements);
        }
        finally
        {
            if (dispose)
            {
                document?.Dispose();
            }
        }
    }

    private int ReplaceInSheet(ExcelDocument document, ExcelSheet sheet)
    {
        var count = 0;
        var range = string.IsNullOrWhiteSpace(Range) ? sheet.GetUsedRangeA1() : Range!;
        using var reader = document.CreateReader();
        var sheetReader = reader.GetSheet(sheet.Name);
        foreach (var cell in sheetReader.EnumerateRange(range))
        {
            if (cell.Value is not string text)
            {
                continue;
            }

            var updated = ReplaceString(text, out var cellReplacements);
            if (cellReplacements == 0)
            {
                continue;
            }

            sheet.CellValue(cell.Row, cell.Column, updated);
            count += cellReplacements;
        }

        return count;
    }

    private string ReplaceString(string value, out int replacements)
    {
        replacements = 0;
        if (Regex.IsPresent)
        {
            var options = CaseSensitive.IsPresent ? RegexOptions.None : RegexOptions.IgnoreCase;
            var count = 0;
            var updated = System.Text.RegularExpressions.Regex.Replace(
                value,
                OldValue,
                match =>
                {
                    count++;
                    return NewValue;
                },
                options);
            replacements = count;
            return updated;
        }

        var comparison = CaseSensitive.IsPresent ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var index = value.IndexOf(OldValue, comparison);
        if (index < 0)
        {
            return value;
        }

        var result = value;
        while (index >= 0)
        {
            result = result.Substring(0, index) + NewValue + result.Substring(index + OldValue.Length);
            replacements++;
            index = result.IndexOf(OldValue, index + NewValue.Length, comparison);
        }

        return result;
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
            return ExcelDocumentService.LoadDocument(resolvedPath, readOnly: false, autoSave: false);
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
}
