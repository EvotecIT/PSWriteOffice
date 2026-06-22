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
    [Parameter(Position = 0)]
    public string? Range { get; set; }

    /// <summary>Header or table column name used to resolve the target range.</summary>
    [Parameter]
    [Alias("ColumnName")]
    public string? HeaderName { get; set; }

    /// <summary>Optional table name for header-based range resolution.</summary>
    [Parameter]
    public string? TableName { get; set; }

    /// <summary>Pivot table name used to resolve the target range.</summary>
    [Parameter]
    public string? PivotTableName { get; set; }

    /// <summary>Use the full pivot output range instead of the default data body range.</summary>
    [Parameter]
    public SwitchParameter PivotWholeTable { get; set; }

    /// <summary>Worksheet header row used when resolving HeaderName without a table. Use 0 for the first row of the used range.</summary>
    [Parameter]
    public int HeaderRow { get; set; }

    /// <summary>Include the header cell in the resolved range.</summary>
    [Parameter]
    public SwitchParameter IncludeHeader { get; set; }

    /// <summary>Rule type to author.</summary>
    [Parameter]
    [Alias("Type")]
    [ValidateSet("CellIs", "Expression", "Formula", "DuplicateValues", "UniqueValues", "Top", "Top10", "Bottom", "Bottom10", "AboveAverage", "BelowAverage", "ContainsText", "NotContainsText", "BeginsWith", "EndsWith", "ContainsBlanks", "NotContainsBlanks", "ContainsErrors", "NotContainsErrors", "TimePeriod")]
    public string RuleType { get; set; } = "CellIs";

    /// <summary>Conditional formatting operator.</summary>
    [Parameter(Position = 1)]
    public string? Operator { get; set; }

    /// <summary>Primary formula or value.</summary>
    [Parameter(Position = 2)]
    public string? Formula1 { get; set; }

    /// <summary>Optional secondary formula or value.</summary>
    [Parameter]
    public string? Formula2 { get; set; }

    /// <summary>Text used by text-matching rule types.</summary>
    [Parameter]
    public string? Text { get; set; }

    /// <summary>Rank used by top/bottom rules.</summary>
    [Parameter]
    public uint Rank { get; set; } = 10;

    /// <summary>Treat top/bottom rank as a percent.</summary>
    [Parameter]
    public SwitchParameter Percent { get; set; }

    /// <summary>Include values equal to the average for average rules.</summary>
    [Parameter]
    public SwitchParameter EqualAverage { get; set; }

    /// <summary>Optional standard deviation threshold for average rules.</summary>
    [Parameter]
    public uint? StandardDeviation { get; set; }

    /// <summary>Time period used by time-period rules.</summary>
    [Parameter]
    public string? TimePeriod { get; set; }

    /// <summary>Stop evaluating later rules when this rule is true.</summary>
    [Parameter]
    public SwitchParameter StopIfTrue { get; set; }

    /// <summary>Emit the range after applying the rule.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();
        string targetRange = ExcelTargetRangeResolver.Resolve(sheet, Range, HeaderName, TableName, HeaderRow, IncludeHeader.IsPresent, PivotTableName, !PivotWholeTable.IsPresent);
        ApplyRule(sheet, targetRange);

        if (PassThru.IsPresent)
        {
            WriteObject(targetRange);
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

    private void ApplyRule(ExcelSheet sheet, string targetRange)
    {
        switch (ResolveRuleType())
        {
            case "CellIs":
                if (string.IsNullOrWhiteSpace(Operator))
                {
                    throw new PSArgumentException("Operator is required for CellIs conditional formatting rules.", nameof(Operator));
                }
                if (string.IsNullOrWhiteSpace(Formula1))
                {
                    throw new PSArgumentException("Formula1 is required for CellIs conditional formatting rules.", nameof(Formula1));
                }
                if (!OpenXmlValueParser.TryParse<ConditionalFormattingOperatorValues>(Operator!, out var op))
                {
                    throw new PSArgumentException($"Unknown conditional formatting operator '{Operator}'.", nameof(Operator));
                }
                sheet.AddConditionalRule(targetRange, op, Formula1!, Formula2);
                break;
            case "Expression":
            case "Formula":
                if (string.IsNullOrWhiteSpace(Formula1))
                {
                    throw new PSArgumentException("Formula1 is required for formula conditional formatting rules.", nameof(Formula1));
                }
                sheet.AddConditionalFormulaRule(targetRange, Formula1!, StopIfTrue.IsPresent);
                break;
            case "DuplicateValues":
                sheet.AddConditionalDuplicateValuesRule(targetRange);
                break;
            case "UniqueValues":
                sheet.AddConditionalUniqueValuesRule(targetRange);
                break;
            case "Top":
            case "Top10":
                sheet.AddConditionalTopBottomRule(targetRange, Rank, bottom: false, percent: Percent.IsPresent);
                break;
            case "Bottom":
            case "Bottom10":
                sheet.AddConditionalTopBottomRule(targetRange, Rank, bottom: true, percent: Percent.IsPresent);
                break;
            case "AboveAverage":
                sheet.AddConditionalAboveAverageRule(targetRange, aboveAverage: true, equalAverage: EqualAverage.IsPresent, standardDeviation: StandardDeviation, stopIfTrue: StopIfTrue.IsPresent);
                break;
            case "BelowAverage":
                sheet.AddConditionalAboveAverageRule(targetRange, aboveAverage: false, equalAverage: EqualAverage.IsPresent, standardDeviation: StandardDeviation, stopIfTrue: StopIfTrue.IsPresent);
                break;
            case "ContainsText":
                sheet.AddConditionalTextRule(targetRange, ConditionalFormatValues.ContainsText, RequireText());
                break;
            case "NotContainsText":
                sheet.AddConditionalTextRule(targetRange, ConditionalFormatValues.NotContainsText, RequireText());
                break;
            case "BeginsWith":
                sheet.AddConditionalTextRule(targetRange, ConditionalFormatValues.BeginsWith, RequireText());
                break;
            case "EndsWith":
                sheet.AddConditionalTextRule(targetRange, ConditionalFormatValues.EndsWith, RequireText());
                break;
            case "ContainsBlanks":
                sheet.AddConditionalBlanksRule(targetRange, containsBlanks: true, stopIfTrue: StopIfTrue.IsPresent);
                break;
            case "NotContainsBlanks":
                sheet.AddConditionalBlanksRule(targetRange, containsBlanks: false, stopIfTrue: StopIfTrue.IsPresent);
                break;
            case "ContainsErrors":
                sheet.AddConditionalErrorsRule(targetRange, containsErrors: true, stopIfTrue: StopIfTrue.IsPresent);
                break;
            case "NotContainsErrors":
                sheet.AddConditionalErrorsRule(targetRange, containsErrors: false, stopIfTrue: StopIfTrue.IsPresent);
                break;
            case "TimePeriod":
                if (string.IsNullOrWhiteSpace(TimePeriod))
                {
                    throw new PSArgumentException("TimePeriod is required for time-period conditional formatting rules.", nameof(TimePeriod));
                }
                if (!OpenXmlValueParser.TryParse<TimePeriodValues>(TimePeriod!, out var period))
                {
                    throw new PSArgumentException($"Unknown conditional formatting time period '{TimePeriod}'.", nameof(TimePeriod));
                }
                sheet.AddConditionalTimePeriodRule(targetRange, period, StopIfTrue.IsPresent);
                break;
            default:
                throw new PSArgumentException($"Unknown conditional formatting rule type '{RuleType}'.", nameof(RuleType));
        }
    }

    private string ResolveRuleType()
    {
        string[] supportedRuleTypes =
        {
            "CellIs", "Expression", "Formula", "DuplicateValues", "UniqueValues",
            "Top", "Top10", "Bottom", "Bottom10", "AboveAverage", "BelowAverage",
            "ContainsText", "NotContainsText", "BeginsWith", "EndsWith",
            "ContainsBlanks", "NotContainsBlanks", "ContainsErrors", "NotContainsErrors",
            "TimePeriod"
        };

        foreach (string supportedRuleType in supportedRuleTypes)
        {
            if (string.Equals(supportedRuleType, RuleType, System.StringComparison.OrdinalIgnoreCase))
            {
                return supportedRuleType;
            }
        }

        return RuleType;
    }

    private string RequireText()
    {
        if (string.IsNullOrWhiteSpace(Text))
        {
            throw new PSArgumentException("Text is required for text conditional formatting rules.", nameof(Text));
        }

        return Text!;
    }
}
