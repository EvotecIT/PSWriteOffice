using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Computes reusable layout boxes for a presentation.</summary>
/// <para>Returns the content box for a slide or equal column/row boxes derived from the current slide size.</para>
/// <example>
///   <summary>Get the content area for a deck.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointLayoutBox -Presentation $ppt -MarginCm 1.5</code>
///   <para>Returns a single layout box representing the usable slide area.</para>
/// </example>
/// <example>
///   <summary>Split the slide into two columns.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointLayoutBox -Presentation $ppt -ColumnCount 2 -MarginCm 1.5 -GutterCm 1.0</code>
///   <para>Returns one layout box per column.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointLayoutBox", DefaultParameterSetName = ParameterSetContent)]
[Alias("PptLayoutBox")]
[OutputType(typeof(PowerPointLayoutBox))]
public sealed class GetOfficePowerPointLayoutBoxCommand : PSCmdlet
{
    private const string ParameterSetContent = "Content";
    private const string ParameterSetColumns = "Columns";
    private const string ParameterSetRows = "Rows";

    /// <summary>Presentation to inspect (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Number of columns to generate.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetColumns)]
    public int ColumnCount { get; set; }

    /// <summary>Number of rows to generate.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetRows)]
    public int RowCount { get; set; }

    /// <summary>Outer slide margin in centimeters.</summary>
    [Parameter]
    public double MarginCm { get; set; }

    /// <summary>Column or row gutter in centimeters.</summary>
    [Parameter(ParameterSetName = ParameterSetColumns)]
    [Parameter(ParameterSetName = ParameterSetRows)]
    public double GutterCm { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            if (MarginCm < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(MarginCm), "MarginCm cannot be negative.");
            }

            if (GutterCm < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(GutterCm), "GutterCm cannot be negative.");
            }

            var presentation = Presentation ?? PowerPointDslContext.Current?.Presentation
                ?? throw new InvalidOperationException("Presentation was not provided. Use -Presentation or run inside New-OfficePowerPoint.");

            switch (ParameterSetName)
            {
                case ParameterSetColumns:
                    foreach (var box in presentation.SlideSize.GetColumnsCm(ColumnCount, MarginCm, GutterCm))
                    {
                        WriteObject(box);
                    }
                    break;
                case ParameterSetRows:
                    foreach (var box in presentation.SlideSize.GetRowsCm(RowCount, MarginCm, GutterCm))
                    {
                        WriteObject(box);
                    }
                    break;
                default:
                    WriteObject(presentation.SlideSize.GetContentBoxCm(MarginCm));
                    break;
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointGetLayoutBoxFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }
}
