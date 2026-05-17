using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Updates OfficeIMO Word table-cell layout and merge settings.</summary>
[Cmdlet(VerbsCommon.Set, "OfficeWordTableCell")]
[Alias("WordTableCellStyle")]
[OutputType(typeof(WordTableCell))]
public sealed class SetOfficeWordTableCellCommand : PSCmdlet
{
    /// <summary>Table cell to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public WordTableCell Cell { get; set; } = null!;

    /// <summary>Cell shading fill color as #RRGGBB.</summary>
    [Parameter] public string? ShadingFillColor { get; set; }

    /// <summary>Cell shading pattern.</summary>
    [Parameter]
    [ValidateSet(
        "Nil",
        "Clear",
        "Solid",
        "HorizontalStripe",
        "VerticalStripe",
        "ReverseDiagonalStripe",
        "DiagonalStripe",
        "HorizontalCross",
        "DiagonalCross",
        "ThinHorizontalStripe",
        "ThinVerticalStripe",
        "ThinReverseDiagonalStripe",
        "ThinDiagonalStripe",
        "ThinHorizontalCross",
        "ThinDiagonalCross",
        "Percent5",
        "Percent10",
        "Percent12",
        "Percent15",
        "Percent20",
        "Percent25",
        "Percent30",
        "Percent35",
        "Percent37",
        "Percent40",
        "Percent45",
        "Percent50",
        "Percent55",
        "Percent60",
        "Percent62",
        "Percent65",
        "Percent70",
        "Percent75",
        "Percent80",
        "Percent85",
        "Percent87",
        "Percent90",
        "Percent95")]
    public string? ShadingPattern { get; set; }

    /// <summary>Cell width value.</summary>
    [Parameter] public int? Width { get; set; }

    /// <summary>Cell width unit type.</summary>
    [Parameter]
    [ValidateSet("Nil", "Pct", "Dxa", "Auto")]
    public string? WidthType { get; set; }

    /// <summary>Cell text direction.</summary>
    [Parameter] public TextDirectionValues? TextDirection { get; set; }

    /// <summary>Whether text wraps in the cell.</summary>
    [Parameter] public bool? WrapText { get; set; }

    /// <summary>Whether text should fit within the cell.</summary>
    [Parameter] public bool? FitText { get; set; }

    /// <summary>Number of cells to merge to the right.</summary>
    [Parameter] public int? MergeRight { get; set; }

    /// <summary>Number of cells to merge downward.</summary>
    [Parameter] public int? MergeDown { get; set; }

    /// <summary>Number of columns to split the cell into.</summary>
    [Parameter] public int? SplitHorizontal { get; set; }

    /// <summary>Number of rows to split the cell into.</summary>
    [Parameter] public int? SplitVertical { get; set; }

    /// <summary>Copy paragraphs while merging cells.</summary>
    [Parameter] public SwitchParameter CopyParagraphs { get; set; }

    /// <summary>Emit the updated table cell.</summary>
    [Parameter] public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Cell == null)
        {
            return;
        }

        if (MyInvocation.BoundParameters.ContainsKey(nameof(ShadingFillColor))) Cell.ShadingFillColorHex = ShadingFillColor ?? string.Empty;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(ShadingPattern))) Cell.ShadingPattern = new ShadingPatternValues(ToOpenXmlToken(ShadingPattern));
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Width))) Cell.Width = Width;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(WidthType))) Cell.WidthType = new TableWidthUnitValues(ToOpenXmlToken(WidthType));
        if (MyInvocation.BoundParameters.ContainsKey(nameof(TextDirection))) Cell.TextDirection = TextDirection;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(WrapText))) Cell.WrapText = WrapText ?? false;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(FitText))) Cell.FitText = FitText ?? false;
        if (MergeRight.HasValue) Cell.MergeHorizontally(MergeRight.Value, CopyParagraphs.IsPresent);
        if (MergeDown.HasValue) Cell.MergeVertically(MergeDown.Value, CopyParagraphs.IsPresent);
        if (SplitHorizontal.HasValue) Cell.SplitHorizontally(SplitHorizontal.Value);
        if (SplitVertical.HasValue) Cell.SplitVertically(SplitVertical.Value);

        if (PassThru.IsPresent)
        {
            WriteObject(Cell);
        }
    }

    private static string ToOpenXmlToken(string? value)
    {
        return value switch
        {
            null => string.Empty,
            "HorizontalStripe" => "horzStripe",
            "VerticalStripe" => "vertStripe",
            "ReverseDiagonalStripe" => "reverseDiagStripe",
            "DiagonalStripe" => "diagStripe",
            "HorizontalCross" => "horzCross",
            "DiagonalCross" => "diagCross",
            "ThinHorizontalStripe" => "thinHorzStripe",
            "ThinVerticalStripe" => "thinVertStripe",
            "ThinReverseDiagonalStripe" => "thinReverseDiagStripe",
            "ThinDiagonalStripe" => "thinDiagStripe",
            "ThinHorizontalCross" => "thinHorzCross",
            "ThinDiagonalCross" => "thinDiagCross",
            "Percent5" => "pct5",
            "Percent10" => "pct10",
            "Percent12" => "pct12",
            "Percent15" => "pct15",
            "Percent20" => "pct20",
            "Percent25" => "pct25",
            "Percent30" => "pct30",
            "Percent35" => "pct35",
            "Percent37" => "pct37",
            "Percent40" => "pct40",
            "Percent45" => "pct45",
            "Percent50" => "pct50",
            "Percent55" => "pct55",
            "Percent60" => "pct60",
            "Percent62" => "pct62",
            "Percent65" => "pct65",
            "Percent70" => "pct70",
            "Percent75" => "pct75",
            "Percent80" => "pct80",
            "Percent85" => "pct85",
            "Percent87" => "pct87",
            "Percent90" => "pct90",
            "Percent95" => "pct95",
            _ => char.ToLowerInvariant(value[0]) + value.Substring(1)
        };
    }
}
