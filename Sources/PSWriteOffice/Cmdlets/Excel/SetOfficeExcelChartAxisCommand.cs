using System;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Configures common Excel chart axis titles, formats, scale, and gridlines.</summary>
/// <example>
///   <summary>Format value axis and gridlines.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$chart | Set-OfficeExcelChartAxis -CategoryTitle 'Month' -ValueTitle 'Revenue' -ValueNumberFormat '$#,##0' -ValueMinimum 0 -ValueMajorUnit 100 -ShowValueMajorGridlines</code>
///   <para>Sets axis titles, value formatting, scale, and major value gridlines.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelChartAxis")]
[Alias("ExcelChartAxis")]
[OutputType(typeof(ExcelChart))]
public sealed class SetOfficeExcelChartAxisCommand : PSCmdlet
{
    /// <summary>Chart to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ExcelChart Chart { get; set; } = null!;

    /// <summary>Axis group to configure.</summary>
    [Parameter]
    public ExcelChartAxisGroup AxisGroup { get; set; } = ExcelChartAxisGroup.Primary;

    /// <summary>Category axis title.</summary>
    [Parameter]
    public string? CategoryTitle { get; set; }

    /// <summary>Value axis title.</summary>
    [Parameter]
    public string? ValueTitle { get; set; }

    /// <summary>Category axis number format code.</summary>
    [Parameter]
    public string? CategoryNumberFormat { get; set; }

    /// <summary>Value axis number format code.</summary>
    [Parameter]
    public string? ValueNumberFormat { get; set; }

    /// <summary>Keep axis number formats linked to source cells.</summary>
    [Parameter]
    public bool SourceLinked { get; set; }

    /// <summary>Value axis minimum.</summary>
    [Parameter]
    public double? ValueMinimum { get; set; }

    /// <summary>Value axis maximum.</summary>
    [Parameter]
    public double? ValueMaximum { get; set; }

    /// <summary>Value axis major unit.</summary>
    [Parameter]
    public double? ValueMajorUnit { get; set; }

    /// <summary>Value axis minor unit.</summary>
    [Parameter]
    public double? ValueMinorUnit { get; set; }

    /// <summary>Show category major gridlines.</summary>
    [Parameter]
    public SwitchParameter ShowCategoryMajorGridlines { get; set; }

    /// <summary>Show category minor gridlines.</summary>
    [Parameter]
    public SwitchParameter ShowCategoryMinorGridlines { get; set; }

    /// <summary>Show value major gridlines.</summary>
    [Parameter]
    public SwitchParameter ShowValueMajorGridlines { get; set; }

    /// <summary>Show value minor gridlines.</summary>
    [Parameter]
    public SwitchParameter ShowValueMinorGridlines { get; set; }

    /// <summary>Optional category gridline color in hex format.</summary>
    [Parameter]
    public string? CategoryGridlineColor { get; set; }

    /// <summary>Optional value gridline color in hex format.</summary>
    [Parameter]
    public string? ValueGridlineColor { get; set; }

    /// <summary>Optional gridline width in points.</summary>
    [Parameter]
    public double? GridlineWidthPoints { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(CategoryTitle))
            {
                Chart.SetCategoryAxisTitle(CategoryTitle, AxisGroup);
            }

            if (!string.IsNullOrWhiteSpace(ValueTitle))
            {
                Chart.SetValueAxisTitle(ValueTitle, AxisGroup);
            }

            if (!string.IsNullOrWhiteSpace(CategoryNumberFormat))
            {
                Chart.SetCategoryAxisNumberFormat(CategoryNumberFormat, SourceLinked, AxisGroup);
            }

            if (!string.IsNullOrWhiteSpace(ValueNumberFormat))
            {
                Chart.SetValueAxisNumberFormat(ValueNumberFormat, SourceLinked, AxisGroup);
            }

            if (ValueMinimum.HasValue || ValueMaximum.HasValue || ValueMajorUnit.HasValue || ValueMinorUnit.HasValue)
            {
                Chart.SetValueAxisScale(ValueMinimum, ValueMaximum, ValueMajorUnit, ValueMinorUnit, axisGroup: AxisGroup);
            }

            bool categoryGridlinesRequested = ShowCategoryMajorGridlines.IsPresent ||
                ShowCategoryMinorGridlines.IsPresent ||
                !string.IsNullOrWhiteSpace(CategoryGridlineColor);
            bool valueGridlinesRequested = ShowValueMajorGridlines.IsPresent ||
                ShowValueMinorGridlines.IsPresent ||
                !string.IsNullOrWhiteSpace(ValueGridlineColor);
            bool widthOnlyRequest = GridlineWidthPoints.HasValue && !categoryGridlinesRequested && !valueGridlinesRequested;

            bool categoryStyleRequested = !string.IsNullOrWhiteSpace(CategoryGridlineColor) ||
                (GridlineWidthPoints.HasValue && (categoryGridlinesRequested || widthOnlyRequest));
            if (categoryGridlinesRequested || widthOnlyRequest)
            {
                bool showMajor = ShowCategoryMajorGridlines.IsPresent || ShowCategoryMinorGridlines.IsPresent || categoryStyleRequested;
                Chart.SetCategoryAxisGridlines(showMajor,
                    ShowCategoryMinorGridlines.IsPresent, CategoryGridlineColor, GridlineWidthPoints, AxisGroup);
            }

            bool valueStyleRequested = !string.IsNullOrWhiteSpace(ValueGridlineColor) ||
                (GridlineWidthPoints.HasValue && (valueGridlinesRequested || widthOnlyRequest));
            if (valueGridlinesRequested || widthOnlyRequest)
            {
                bool showMajor = ShowValueMajorGridlines.IsPresent || ShowValueMinorGridlines.IsPresent || valueStyleRequested;
                Chart.SetValueAxisGridlines(showMajor,
                    ShowValueMinorGridlines.IsPresent, ValueGridlineColor, GridlineWidthPoints, AxisGroup);
            }

            WriteObject(Chart);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ExcelChartAxisFailed", ErrorCategory.InvalidOperation, Chart));
        }
    }
}
