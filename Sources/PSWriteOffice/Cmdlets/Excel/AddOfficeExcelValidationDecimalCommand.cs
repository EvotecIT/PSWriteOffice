using System.Management.Automation;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a decimal-number data validation rule to a worksheet range.</summary>
/// <example>
///   <summary>Restrict values between 0 and 1.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelValidationDecimal -Range 'B2:B20' -Operator Between -Formula1 0 -Formula2 1 }</code>
///   <para>Ensures values in B2:B20 are decimals between 0 and 1.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelValidationDecimal", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelValidationDecimal")]
public sealed class AddOfficeExcelValidationDecimalCommand : PSCmdlet
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

    /// <summary>Validation operator.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Operator { get; set; } = string.Empty;

    /// <summary>Primary numeric bound.</summary>
    [Parameter(Mandatory = true, Position = 2)]
    public double Formula1 { get; set; }

    /// <summary>Optional secondary numeric bound.</summary>
    [Parameter]
    public double? Formula2 { get; set; }

    /// <summary>Allow blank values.</summary>
    [Parameter]
    public bool AllowBlank { get; set; } = true;

    /// <summary>Error title shown to the user.</summary>
    [Parameter]
    public string? ErrorTitle { get; set; }

    /// <summary>Error message shown to the user.</summary>
    [Parameter]
    public string? ErrorMessage { get; set; }

    /// <summary>Emit the range string after applying validation.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (!OpenXmlValueParser.TryParse<DataValidationOperatorValues>(Operator, out var op))
        {
            throw new PSArgumentException($"Unknown validation operator '{Operator}'.", nameof(Operator));
        }

        var sheet = ResolveSheet();
        sheet.ValidationDecimal(Range, op, Formula1, Formula2, AllowBlank, ErrorTitle, ErrorMessage);

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

            return ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex);
        }

        var context = ExcelDslContext.Require(this);
        return context.RequireSheet();
    }
}
