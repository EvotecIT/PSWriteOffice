using System;
using System.Management.Automation;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Configures legend visibility and styling for an Excel chart.</summary>
/// <example>
///   <summary>Move the legend to the right.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$chart | Set-OfficeExcelChartLegend -Position Right</code>
///   <para>Shows the legend on the right side of the chart and returns the chart for chaining.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelChartLegend")]
[OutputType(typeof(ExcelChart))]
public sealed class SetOfficeExcelChartLegendCommand : PSCmdlet
{
    /// <summary>Chart to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ExcelChart Chart { get; set; } = null!;

    /// <summary>Legend position.</summary>
    [Parameter]
    [ValidateSet("Bottom", "Left", "Right", "Top", "TopRight")]
    public string Position { get; set; } = "Right";

    /// <summary>Overlay the legend on the chart area.</summary>
    [Parameter]
    public bool Overlay { get; set; }

    /// <summary>Hide the legend instead of positioning it.</summary>
    [Parameter]
    public SwitchParameter Hide { get; set; }

    /// <summary>Optional legend font size in points.</summary>
    [Parameter]
    public double? FontSizePoints { get; set; }

    /// <summary>Optional bold setting for legend text.</summary>
    [Parameter]
    public bool? Bold { get; set; }

    /// <summary>Optional italic setting for legend text.</summary>
    [Parameter]
    public bool? Italic { get; set; }

    /// <summary>Optional legend text color in hex format.</summary>
    [Parameter]
    public string? Color { get; set; }

    /// <summary>Optional legend font name.</summary>
    [Parameter]
    public string? FontName { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            if (Hide.IsPresent)
            {
                Chart.HideLegend();
            }
            else
            {
                Chart.SetLegend(ResolveLegendPosition(Position), Overlay);
            }

            if (FontSizePoints.HasValue || Bold.HasValue || Italic.HasValue ||
                !string.IsNullOrWhiteSpace(Color) || !string.IsNullOrWhiteSpace(FontName))
            {
                Chart.SetLegendTextStyle(FontSizePoints, Bold, Italic, Color, FontName);
            }

            WriteObject(Chart);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ExcelChartLegendFailed", ErrorCategory.InvalidOperation, Chart));
        }
    }

    private static C.LegendPositionValues ResolveLegendPosition(string value)
    {
        return value switch
        {
            "Bottom" => C.LegendPositionValues.Bottom,
            "Left" => C.LegendPositionValues.Left,
            "Right" => C.LegendPositionValues.Right,
            "Top" => C.LegendPositionValues.Top,
            "TopRight" => C.LegendPositionValues.TopRight,
            _ => throw new PSArgumentException($"Unsupported legend position '{value}'.")
        };
    }
}
