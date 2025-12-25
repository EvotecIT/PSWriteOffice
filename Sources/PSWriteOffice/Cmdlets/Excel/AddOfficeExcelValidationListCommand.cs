using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a list-based data validation to a worksheet range.</summary>
/// <para>When invoked inside the Excel DSL, applies to the current worksheet.</para>
/// <example>
///   <summary>Add a validation list to a range.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelValidationList -Range 'C2:C50' -Values 'New','In Progress','Done' }</code>
///   <para>Restricts column C to the provided values.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelValidationList", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelValidationList")]
public sealed class AddOfficeExcelValidationListCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public int? SheetIndex { get; set; }

    /// <summary>Target range in A1 notation.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Range { get; set; } = string.Empty;

    /// <summary>Allowed values for the dropdown list.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string[] Values { get; set; } = Array.Empty<string>();

    /// <summary>Allow blank values.</summary>
    [Parameter]
    public bool AllowBlank { get; set; } = true;

    /// <summary>Emit the range string after applying validation.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Values == null || Values.Length == 0)
        {
            throw new PSArgumentException("Provide at least one validation value.");
        }

        var sheet = ResolveSheet();
        sheet.ValidationList(Range, Values, AllowBlank);

        if (PassThru.IsPresent)
        {
            WriteObject(Range);
        }
    }

    private ExcelSheet ResolveSheet()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            if (Document == null)
            {
                throw new PSArgumentException("Provide an Excel document.");
            }

            if (!string.IsNullOrWhiteSpace(Sheet))
            {
                return Document[Sheet!];
            }

            if (SheetIndex.HasValue)
            {
                if (SheetIndex.Value < 0 || SheetIndex.Value >= Document.Sheets.Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(SheetIndex), "SheetIndex is out of range.");
                }
                return Document.Sheets[SheetIndex.Value];
            }

            if (Document.Sheets.Count == 0)
            {
                throw new InvalidOperationException("Workbook contains no worksheets.");
            }

            return Document.Sheets[0];
        }

        var context = ExcelDslContext.Require(this);
        return context.RequireSheet();
    }
}
