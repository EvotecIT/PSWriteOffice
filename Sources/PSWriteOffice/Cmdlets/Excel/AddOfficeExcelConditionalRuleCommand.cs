using System.Management.Automation;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a conditional formatting rule to the current worksheet.</summary>
/// <example>
///   <summary>Highlight values greater than 100.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelConditionalRule -Range 'C2:C100' -Operator GreaterThan -Formula1 '100' }</code>
///   <para>Applies a conditional rule to column C.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelConditionalRule", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelConditionalRule")]
public sealed class AddOfficeExcelConditionalRuleCommand : PSCmdlet
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

    /// <summary>A1 range to apply the rule to.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Range { get; set; } = string.Empty;

    /// <summary>Conditional formatting operator.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Operator { get; set; } = string.Empty;

    /// <summary>Primary formula or value.</summary>
    [Parameter(Mandatory = true, Position = 2)]
    public string Formula1 { get; set; } = string.Empty;

    /// <summary>Optional secondary formula or value.</summary>
    [Parameter]
    public string? Formula2 { get; set; }

    /// <summary>Emit the range after applying the rule.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (!OpenXmlValueParser.TryParse<ConditionalFormattingOperatorValues>(Operator, out var op))
        {
            throw new PSArgumentException($"Unknown conditional formatting operator '{Operator}'.", nameof(Operator));
        }

        var sheet = ResolveSheet();
        sheet.AddConditionalRule(Range, op, Formula1, Formula2);

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
