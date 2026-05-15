using System;
using System.Management.Automation;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Excel;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Configures Excel chart series colors, line style, and markers.</summary>
/// <example>
///   <summary>Format a series by name.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$chart | Set-OfficeExcelChartSeries -SeriesName 'Revenue' -FillColor '#4472C4' -LineColor '#1F4E79' -MarkerStyle Circle -MarkerSize 6</code>
///   <para>Applies fill, line, and marker settings to the Revenue series.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelChartSeries", DefaultParameterSetName = ParameterSetIndex)]
[Alias("ExcelChartSeries")]
[OutputType(typeof(ExcelChart))]
public sealed class SetOfficeExcelChartSeriesCommand : PSCmdlet
{
    private const string ParameterSetIndex = "Index";
    private const string ParameterSetSeriesName = "Name";

    /// <summary>Chart to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ExcelChart Chart { get; set; } = null!;

    /// <summary>Zero-based series index.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetIndex)]
    public int SeriesIndex { get; set; }

    /// <summary>Series name.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetSeriesName)]
    public string SeriesName { get; set; } = string.Empty;

    /// <summary>Ignore case when matching series name.</summary>
    [Parameter(ParameterSetName = ParameterSetSeriesName)]
    public bool IgnoreCase { get; set; } = true;

    /// <summary>Series fill color in hex format.</summary>
    [Parameter]
    public string? FillColor { get; set; }

    /// <summary>Series line color in hex format.</summary>
    [Parameter]
    public string? LineColor { get; set; }

    /// <summary>Series line width in points.</summary>
    [Parameter]
    public double? LineWidthPoints { get; set; }

    /// <summary>Marker style name.</summary>
    [Parameter]
    public string? MarkerStyle { get; set; }

    /// <summary>Marker size.</summary>
    [Parameter]
    public int? MarkerSize { get; set; }

    /// <summary>Marker fill color in hex format.</summary>
    [Parameter]
    public string? MarkerFillColor { get; set; }

    /// <summary>Marker line color in hex format.</summary>
    [Parameter]
    public string? MarkerLineColor { get; set; }

    /// <summary>Marker line width in points.</summary>
    [Parameter]
    public double? MarkerLineWidthPoints { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(FillColor))
            {
                if (ParameterSetName == ParameterSetSeriesName)
                {
                    Chart.SetSeriesFillColor(SeriesName, FillColor, IgnoreCase);
                }
                else
                {
                    Chart.SetSeriesFillColor(SeriesIndex, FillColor);
                }
            }

            if (!string.IsNullOrWhiteSpace(LineColor) || LineWidthPoints.HasValue)
            {
                if (ParameterSetName == ParameterSetSeriesName)
                {
                    Chart.SetSeriesLineColor(SeriesName, LineColor ?? "000000", LineWidthPoints, IgnoreCase);
                }
                else
                {
                    Chart.SetSeriesLineColor(SeriesIndex, LineColor ?? "000000", LineWidthPoints);
                }
            }

            bool markerRequested = !string.IsNullOrWhiteSpace(MarkerStyle) || MarkerSize.HasValue ||
                !string.IsNullOrWhiteSpace(MarkerFillColor) || !string.IsNullOrWhiteSpace(MarkerLineColor) || MarkerLineWidthPoints.HasValue;
            if (markerRequested)
            {
                string markerStyleName = string.IsNullOrWhiteSpace(MarkerStyle) ? "Circle" : MarkerStyle;
                if (!OpenXmlValueParser.TryParse(markerStyleName, out C.MarkerStyleValues markerStyle))
                {
                    throw new PSArgumentException($"Unknown MarkerStyle '{MarkerStyle}'.");
                }

                if (ParameterSetName == ParameterSetSeriesName)
                {
                    Chart.SetSeriesMarker(SeriesName, markerStyle, MarkerSize, MarkerFillColor, MarkerLineColor, MarkerLineWidthPoints, IgnoreCase);
                }
                else
                {
                    Chart.SetSeriesMarker(SeriesIndex, markerStyle, MarkerSize, MarkerFillColor, MarkerLineColor, MarkerLineWidthPoints);
                }
            }

            WriteObject(Chart);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ExcelChartSeriesFailed", ErrorCategory.InvalidOperation, Chart));
        }
    }
}
