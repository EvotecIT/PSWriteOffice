using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Chart types supported by Add-OfficePowerPointChart.</summary>
public enum PowerPointChartType
{
    /// <summary>Clustered column chart.</summary>
    ClusteredColumn,
    /// <summary>Line chart.</summary>
    Line,
    /// <summary>Pie chart.</summary>
    Pie,
    /// <summary>Doughnut chart.</summary>
    Doughnut,
    /// <summary>Scatter chart.</summary>
    Scatter
}

/// <summary>Adds a chart to a PowerPoint slide.</summary>
/// <para>Supports default chart data or object-based category/series mappings for standard and scatter charts.</para>
/// <example>
///   <summary>Add a clustered column chart from objects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficePowerPointChart -Slide $slide -Data $rows -CategoryProperty Month -SeriesProperty Sales,Profit -Title 'Monthly performance'</code>
///   <para>Creates a clustered column chart using Month for categories and Sales/Profit as series.</para>
/// </example>
/// <example>
///   <summary>Add a scatter chart from numeric properties.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficePowerPointChart -Slide $slide -Type Scatter -Data $rows -XProperty Quarter -YProperty Revenue -Title 'Revenue trend'</code>
///   <para>Creates a scatter chart using Quarter on the X axis and Revenue on the Y axis.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointChart", DefaultParameterSetName = ParameterSetDefault)]
[Alias("PptChart")]
[OutputType(typeof(PowerPointChart))]
public sealed class AddOfficePowerPointChartCommand : PSCmdlet
{
    private const string ParameterSetDefault = "Default";
    private const string ParameterSetCategorical = "Categorical";
    private const string ParameterSetScatter = "Scatter";

    /// <summary>Target slide that will receive the chart (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Chart type to create.</summary>
    [Parameter]
    public PowerPointChartType Type { get; set; } = PowerPointChartType.ClusteredColumn;

    /// <summary>Source objects used to build chart data.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetCategorical)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetScatter)]
    public object[] Data { get; set; } = Array.Empty<object>();

    /// <summary>Property name used for category labels on standard charts.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetCategorical)]
    public string CategoryProperty { get; set; } = string.Empty;

    /// <summary>Property names used as numeric series on standard charts.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetCategorical)]
    public string[] SeriesProperty { get; set; } = Array.Empty<string>();

    /// <summary>Property name used for the X axis on scatter charts.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetScatter)]
    public string XProperty { get; set; } = string.Empty;

    /// <summary>Property names used as numeric Y series on scatter charts.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetScatter)]
    public string[] YProperty { get; set; } = Array.Empty<string>();

    /// <summary>Left offset in points from the slide origin.</summary>
    [Parameter]
    public double X { get; set; } = 40;

    /// <summary>Top offset in points from the slide origin.</summary>
    [Parameter]
    public double Y { get; set; } = 120;

    /// <summary>Chart width in points.</summary>
    [Parameter]
    public double Width { get; set; } = 420;

    /// <summary>Chart height in points.</summary>
    [Parameter]
    public double Height { get; set; } = 240;

    /// <summary>Optional chart title.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            if (Width <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(Width), "Width must be greater than 0.");
            }

            if (Height <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(Height), "Height must be greater than 0.");
            }

            var slide = Slide ?? PowerPointDslContext.Require(this).RequireSlide();
            var chart = ParameterSetName switch
            {
                ParameterSetCategorical => AddCategoricalChart(slide),
                ParameterSetScatter => AddScatterChart(slide),
                _ => AddDefaultChart(slide)
            };

            if (!string.IsNullOrWhiteSpace(Title))
            {
                chart.SetTitle(Title!);
            }

            WriteObject(chart);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointAddChartFailed", ErrorCategory.InvalidOperation, Slide));
        }
    }

    private PowerPointChart AddDefaultChart(PowerPointSlide slide)
    {
        return Type switch
        {
            PowerPointChartType.Line => slide.AddLineChartPoints(X, Y, Width, Height),
            PowerPointChartType.Pie => slide.AddPieChartPoints(X, Y, Width, Height),
            PowerPointChartType.Doughnut => slide.AddDoughnutChartPoints(X, Y, Width, Height),
            PowerPointChartType.Scatter => slide.AddScatterChartPoints(X, Y, Width, Height),
            _ => slide.AddChartPoints(X, Y, Width, Height)
        };
    }

    private PowerPointChart AddCategoricalChart(PowerPointSlide slide)
    {
        if (Type == PowerPointChartType.Scatter)
        {
            throw new PSArgumentException("Use -XProperty and -YProperty when -Type Scatter is selected.", nameof(Type));
        }

        if (SeriesProperty == null || SeriesProperty.Length == 0)
        {
            throw new PSArgumentException("Provide at least one -SeriesProperty.", nameof(SeriesProperty));
        }

        if ((Type == PowerPointChartType.Pie || Type == PowerPointChartType.Doughnut) && SeriesProperty.Length > 1)
        {
            throw new PSArgumentException("Pie and Doughnut charts support only one series property.", nameof(SeriesProperty));
        }

        var data = BuildChartData();
        return Type switch
        {
            PowerPointChartType.Line => slide.AddLineChartPoints(data, X, Y, Width, Height),
            PowerPointChartType.Pie => slide.AddPieChartPoints(data, X, Y, Width, Height),
            PowerPointChartType.Doughnut => slide.AddDoughnutChartPoints(data, X, Y, Width, Height),
            _ => slide.AddChartPoints(data, X, Y, Width, Height)
        };
    }

    private PowerPointChart AddScatterChart(PowerPointSlide slide)
    {
        if (Type != PowerPointChartType.Scatter)
        {
            throw new PSArgumentException("Use -CategoryProperty/-SeriesProperty for non-scatter charts.", nameof(Type));
        }

        if (YProperty == null || YProperty.Length == 0)
        {
            throw new PSArgumentException("Provide at least one -YProperty.", nameof(YProperty));
        }

        var data = BuildScatterChartData();
        return slide.AddScatterChartPoints(data, X, Y, Width, Height);
    }

    private PowerPointChartData BuildChartData()
    {
        var items = EnsureData();
        var categories = items.Select(item => ConvertToString(GetPropertyValue(item, CategoryProperty), CategoryProperty)).ToArray();
        var series = SeriesProperty.Select(property =>
            new PowerPointChartSeries(property, items.Select(item => ConvertToDouble(GetPropertyValue(item, property), property)).ToArray()))
            .ToArray();

        return new PowerPointChartData(categories, series);
    }

    private PowerPointScatterChartData BuildScatterChartData()
    {
        var items = EnsureData();
        var xValues = items.Select(item => ConvertToDouble(GetPropertyValue(item, XProperty), XProperty)).ToArray();
        var series = YProperty.Select(property =>
            new PowerPointScatterChartSeries(
                property,
                xValues,
                items.Select(item => ConvertToDouble(GetPropertyValue(item, property), property)).ToArray()))
            .ToArray();

        return new PowerPointScatterChartData(series);
    }

    private object[] EnsureData()
    {
        if (Data == null || Data.Length == 0)
        {
            throw new PSArgumentException("Provide at least one data item.", nameof(Data));
        }

        return Data;
    }

    private static object? GetPropertyValue(object item, string propertyName)
    {
        if (item == null)
        {
            throw new PSArgumentException("Chart data items cannot be null.");
        }

        if (string.IsNullOrWhiteSpace(propertyName))
        {
            throw new PSArgumentException("Property name cannot be empty.", nameof(propertyName));
        }

        if (item is IDictionary dictionary)
        {
            foreach (DictionaryEntry entry in dictionary)
            {
                if (entry.Key is string key && string.Equals(key, propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    return entry.Value;
                }
            }
        }

        var property = PSObject.AsPSObject(item).Properties[propertyName];
        if (property == null)
        {
            throw new PSArgumentException($"Property '{propertyName}' was not found on chart data item.");
        }

        return property.Value;
    }

    private static string ConvertToString(object? value, string propertyName)
    {
        if (value == null)
        {
            throw new PSArgumentException($"Property '{propertyName}' cannot be null.");
        }

        var text = Convert.ToString(value, CultureInfo.InvariantCulture);
        if (string.IsNullOrWhiteSpace(text))
        {
            throw new PSArgumentException($"Property '{propertyName}' cannot be empty.");
        }

        return text;
    }

    private static double ConvertToDouble(object? value, string propertyName)
    {
        if (value == null)
        {
            throw new PSArgumentException($"Property '{propertyName}' cannot be null.");
        }

        try
        {
            return Convert.ToDouble(value, CultureInfo.InvariantCulture);
        }
        catch (Exception)
        {
            throw new PSArgumentException($"Property '{propertyName}' must be numeric.", propertyName);
        }
    }
}
