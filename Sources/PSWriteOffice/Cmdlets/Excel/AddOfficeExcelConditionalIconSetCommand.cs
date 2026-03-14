using System;
using System.Management.Automation;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds an icon set conditional format to a range.</summary>
/// <example>
///   <summary>Add traffic light icons.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelConditionalIconSet -Range 'E2:E50' -IconSet ThreeTrafficLights1 }</code>
///   <para>Applies a traffic-light icon set.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelConditionalIconSet", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelConditionalIconSet")]
public sealed class AddOfficeExcelConditionalIconSetCommand : PSCmdlet
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

    /// <summary>A1 range to format.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Range { get; set; } = string.Empty;

    /// <summary>Icon set to apply.</summary>
    [Parameter]
    public string IconSet { get; set; } = "ThreeTrafficLights1";

    /// <summary>Show the underlying values.</summary>
    [Parameter]
    public bool ShowValue { get; set; } = true;

    /// <summary>Reverse the icon order.</summary>
    [Parameter]
    public bool Reverse { get; set; }

    /// <summary>Percent thresholds (0..100) matching the icon count.</summary>
    [Parameter]
    public double[]? PercentThresholds { get; set; }

    /// <summary>Number thresholds matching the icon count.</summary>
    [Parameter]
    public double[]? NumberThresholds { get; set; }

    /// <summary>Emit the range after applying the format.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (PercentThresholds != null && NumberThresholds != null)
        {
            throw new PSArgumentException("Provide either PercentThresholds or NumberThresholds, not both.");
        }

        if (!OpenXmlValueParser.TryParse<IconSetValues>(IconSet, out var iconSet))
        {
            throw new PSArgumentException($"Unknown icon set '{IconSet}'.", nameof(IconSet));
        }

        var sheet = ResolveSheet();
        sheet.AddConditionalIconSet(
            Range,
            iconSet,
            showValue: ShowValue,
            reverseIconOrder: Reverse,
            percentThresholds: PercentThresholds,
            numberThresholds: NumberThresholds);

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
