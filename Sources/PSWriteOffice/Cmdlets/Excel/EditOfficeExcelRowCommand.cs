using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Runs a script block against editable worksheet rows.</summary>
/// <example>
///   <summary>Edit rows by header name.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Edit-OfficeExcelRow -Path .\Report.xlsx -Sheet Data -ScriptBlock { param($row) if ($row.Get[string]('Status') -eq 'Draft') { $row.Set('Status', 'Ready') } }</code>
///   <para>Loads editable row handles, lets the script update cells, and saves the workbook.</para>
/// </example>
[Cmdlet(VerbsData.Edit, "OfficeExcelRow", DefaultParameterSetName = ParameterSetPath)]
[Alias("Edit-ExcelRow", "ExcelRowEdit")]
[OutputType(typeof(RowEdit))]
public sealed class EditOfficeExcelRowCommand : PSCmdlet
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

    /// <summary>Worksheet name. Defaults to the current sheet inside an ExcelSheet block.</summary>
    [Parameter]
    [Alias("WorksheetName")]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index when using a workbook object or path.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>A1 range to expose as editable rows. Defaults to the worksheet used range.</summary>
    [Parameter]
    public string? Range { get; set; }

    /// <summary>Script block to run once per editable row. The row is passed as the first argument.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public ScriptBlock ScriptBlock { get; set; } = null!;

    /// <summary>Prefer decimals instead of doubles for numeric values.</summary>
    [Parameter]
    public SwitchParameter NumericAsDecimal { get; set; }

    /// <summary>Emit each editable row after the script block runs.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDocument? document = null;
        var dispose = false;

        try
        {
            document = ResolveDocument(out dispose);
            var sheet = ResolveSheet(document);
            var options = ExcelReadOutputService.CreateOptions(NumericAsDecimal.IsPresent);
            var rows = string.IsNullOrWhiteSpace(Range)
                ? sheet.RowsObjects(options)
                : sheet.RowsObjects(Range!, options);

            foreach (var row in rows)
            {
                ScriptBlock.Invoke(row);
                if (PassThru.IsPresent)
                {
                    WriteObject(row);
                }
            }

            if (dispose)
            {
                document.Save(false);
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
            return ExcelDocumentService.LoadDocument(resolvedPath, readOnly: false, autoSave: false);
        }

        return ParameterSetName == ParameterSetDocument
            ? Document ?? throw new PSArgumentException("Provide an Excel document.")
            : ExcelDslContext.Require(this).Document;
    }

    private ExcelSheet ResolveSheet(ExcelDocument document)
    {
        if (ParameterSetName == ParameterSetContext && string.IsNullOrWhiteSpace(Sheet) && !SheetIndex.HasValue)
        {
            return ExcelDslContext.Require(this).RequireSheet();
        }

        return ExcelSheetResolver.Resolve(document, Sheet, SheetIndex);
    }
}
