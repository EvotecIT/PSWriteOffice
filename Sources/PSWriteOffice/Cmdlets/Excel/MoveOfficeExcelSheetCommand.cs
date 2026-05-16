using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Moves a worksheet to a new workbook position.</summary>
/// <example>
///   <summary>Move a sheet to the front.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Move-OfficeExcelSheet -Path .\Report.xlsx -Sheet Summary -Index 0</code>
///   <para>Moves Summary to the first worksheet tab.</para>
/// </example>
[Cmdlet(VerbsCommon.Move, "OfficeExcelSheet", DefaultParameterSetName = ParameterSetContext)]
[Alias("Set-OfficeExcelSheetOrder", "ExcelSheetOrder")]
public sealed class MoveOfficeExcelSheetCommand : PSCmdlet
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

    /// <summary>Worksheet name to move. Defaults to the current sheet inside an ExcelSheet block.</summary>
    [Parameter]
    [Alias("WorksheetName")]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index to move when using a workbook object or path.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>Zero-based destination tab index.</summary>
    [Parameter(Mandatory = true)]
    [Alias("TargetIndex")]
    public int Index { get; set; }

    /// <summary>Emit the moved worksheet.</summary>
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
            document.ReorderWorksheet(sheet, Index);

            if (dispose)
            {
                document.Save(false);
            }

            if (PassThru.IsPresent)
            {
                WriteObject(sheet);
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
