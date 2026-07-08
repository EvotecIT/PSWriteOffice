using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a table to a PDF document.</summary>
/// <example>
///   <summary>Add object data as a PDF table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$services = @(
///     [pscustomobject]@{ Name = 'Directory'; Status = 'Healthy'; Incidents = 0 }
///     [pscustomobject]@{ Name = 'Mail'; Status = 'Watch'; Incidents = 2 }
/// )
/// New-OfficePdf -Path .\Examples\Documents\PdfTable.pdf {
///     Add-OfficePdfHeading -Text 'Service status'
///     Add-OfficePdfTable -InputObject $services -Property Name,Status,Incidents -Header 'Service','Status','Incidents'
/// }</code>
///   <para>Converts PowerShell objects into a table using selected properties and friendly headers.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfTable", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfTable")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfTableCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPipelineDocument = "PipelineDocument";
    private readonly List<object?> _items = new();

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocument)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetPipelineDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Objects or row arrays to render as a table.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetContext)]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPipelineDocument)]
    public object? InputObject { get; set; }

    /// <summary>Specific object properties to include.</summary>
    [Parameter]
    public string[]? Property { get; set; }

    /// <summary>Header labels. Defaults to property names.</summary>
    [Parameter]
    public string[]? Header { get; set; }

    /// <summary>Projection to apply before writing the table.</summary>
    [Parameter]
    public OfficeTableView View { get; set; } = OfficeTableView.Normal;

    /// <summary>Table alignment.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Left;

    /// <summary>OfficeIMO table style preset or supported Word table style name.</summary>
    [Parameter]
    public string? TableStyle { get; set; }

    /// <summary>Header fill color in #RRGGBB format.</summary>
    [Parameter]
    public string? HeaderFill { get; set; }

    /// <summary>Header text color in #RRGGBB format.</summary>
    [Parameter]
    public string? HeaderTextColor { get; set; }

    /// <summary>Body text color in #RRGGBB format.</summary>
    [Parameter]
    public string? TextColor { get; set; }

    /// <summary>Alternating body row fill color in #RRGGBB format.</summary>
    [Parameter]
    public string? RowStripeFill { get; set; }

    /// <summary>Border color in #RRGGBB format.</summary>
    [Parameter]
    public string? BorderColor { get; set; }

    /// <summary>Border width in PDF points.</summary>
    [Parameter]
    public double? BorderWidth { get; set; }

    /// <summary>Body cell font size in PDF points.</summary>
    [Parameter]
    public double? FontSize { get; set; }

    /// <summary>Header cell font size in PDF points.</summary>
    [Parameter]
    public double? HeaderFontSize { get; set; }

    /// <summary>Wrapped line height multiplier for table cells.</summary>
    [Parameter]
    public double? LineHeight { get; set; }

    /// <summary>Horizontal cell padding in PDF points.</summary>
    [Parameter]
    public double? CellPaddingX { get; set; }

    /// <summary>Vertical cell padding in PDF points.</summary>
    [Parameter]
    public double? CellPaddingY { get; set; }

    /// <summary>Spacing before the table in PDF points.</summary>
    [Parameter]
    public double? SpacingBefore { get; set; }

    /// <summary>Spacing after the table in PDF points.</summary>
    [Parameter]
    public double? SpacingAfter { get; set; }

    /// <summary>Caption rendered above the table grid.</summary>
    [Parameter]
    public string? Caption { get; set; }

    /// <summary>Caption alignment.</summary>
    [Parameter]
    public PdfAlign? CaptionAlign { get; set; }

    /// <summary>Caption color in #RRGGBB format.</summary>
    [Parameter]
    public string? CaptionColor { get; set; }

    /// <summary>Caption font size in PDF points.</summary>
    [Parameter]
    public double? CaptionFontSize { get; set; }

    /// <summary>Fixed column widths in PDF points.</summary>
    [Parameter]
    public double[]? ColumnWidthPoints { get; set; }

    /// <summary>Relative column width weights.</summary>
    [Parameter]
    public double[]? ColumnWidthWeights { get; set; }

    /// <summary>Per-column horizontal alignment.</summary>
    [Parameter]
    public PdfColumnAlign[]? ColumnAlign { get; set; }

    /// <summary>Measure flexible columns from content.</summary>
    [Parameter]
    public SwitchParameter AutoFitColumns { get; set; }

    /// <summary>Right-align numeric-looking cell values.</summary>
    [Parameter]
    public SwitchParameter RightAlignNumeric { get; set; }

    /// <summary>Reduce table text size when needed so cell text fits within the resolved cell width.</summary>
    [Parameter]
    public SwitchParameter ShrinkTextToFit { get; set; }

    /// <summary>Smallest font size, in points, used by -ShrinkTextToFit.</summary>
    [Parameter]
    public double? MinimumShrinkFontSize { get; set; }

    /// <summary>Keep the table together when possible.</summary>
    [Parameter]
    public SwitchParameter KeepTogether { get; set; }

    /// <summary>Keep the table with the next block when possible.</summary>
    [Parameter]
    public SwitchParameter KeepWithNext { get; set; }

    /// <summary>Hide table borders.</summary>
    [Parameter]
    public SwitchParameter NoBorder { get; set; }

    /// <summary>Disable the header fill.</summary>
    [Parameter]
    public SwitchParameter NoHeaderFill { get; set; }

    /// <summary>Disable alternating row fill.</summary>
    [Parameter]
    public SwitchParameter NoRowStripeFill { get; set; }

    /// <summary>Number of leading rows rendered as header rows.</summary>
    [Parameter]
    public int? HeaderRowCount { get; set; }

    /// <summary>Number of leading header rows repeated on following pages.</summary>
    [Parameter]
    public int? RepeatHeaderRowCount { get; set; }

    /// <summary>Number of trailing rows rendered as footer rows.</summary>
    [Parameter]
    public int? FooterRowCount { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (ParameterSetName == ParameterSetPipelineDocument)
        {
            RenderTable(Document, BuildRows(InputObject));
            if (PassThru.IsPresent)
            {
                WriteObject(Document);
            }

            return;
        }

        TableInputCollector.AddInput(_items, InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        if (ParameterSetName == ParameterSetPipelineDocument)
        {
            return;
        }

        var document = PdfCommandUtilities.ResolveDocument(this, Document, ParameterSetName, ParameterSetDocument);
        RenderTable(document, TableInputCollector.RequireRows(_items, nameof(InputObject)));
        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }

    private void RenderTable(PdfDocument document, object[] inputRows)
    {
        var projectedRows = TableViewProjection.Project(inputRows, View);
        if (OfficeTableSpecParser.TryCreate(projectedRows, Property, Header, out var tableSpec))
        {
            document.Table(ToPdfRows(tableSpec), Align, CreateStyle());
            return;
        }

        var rowArrayInput = projectedRows.All(item => item is IEnumerable && item is not string && item is not IDictionary);
        string[][] rows = rowArrayInput
            ? PdfCommandUtilities.ConvertDataRows(projectedRows, Header)
            : projectedRows.Length == 1 && projectedRows[0] is IEnumerable enumerable && projectedRows[0] is not string && projectedRows[0] is not IDictionary
            ? PdfCommandUtilities.ConvertDataRows(enumerable, Header)
            : PdfCommandUtilities.ConvertToTableRows(projectedRows, Property, Header);

        document.Table(rows, Align, CreateStyle());
    }

    private static PdfTableCell[][] ToPdfRows(OfficeTableSpec table)
    {
        return table.Rows
            .Select(row => row
                .Select(cell => cell.HasSpan
                    ? PdfTableCell.Merge(cell.Text, cell.ColumnSpan, cell.RowSpan)
                    : PdfTableCell.TextCell(cell.Text))
                .ToArray())
            .ToArray();
    }

    private PdfTableStyle? CreateStyle()
    {
        return PdfTableStyleBuilder.Create(new PdfTableStyleOptions
        {
            TableStyle = TableStyle,
            HeaderFill = HeaderFill,
            HeaderTextColor = HeaderTextColor,
            TextColor = TextColor,
            RowStripeFill = RowStripeFill,
            BorderColor = BorderColor,
            BorderWidth = BorderWidth,
            FontSize = FontSize,
            HeaderFontSize = HeaderFontSize,
            LineHeight = LineHeight,
            CellPaddingX = CellPaddingX,
            CellPaddingY = CellPaddingY,
            SpacingBefore = SpacingBefore,
            SpacingAfter = SpacingAfter,
            Caption = Caption,
            CaptionAlign = CaptionAlign,
            CaptionColor = CaptionColor,
            CaptionFontSize = CaptionFontSize,
            ColumnWidthPoints = ColumnWidthPoints,
            ColumnWidthWeights = ColumnWidthWeights,
            ColumnAlign = ColumnAlign,
            AutoFitColumns = AutoFitColumns.IsPresent,
            RightAlignNumeric = RightAlignNumeric.IsPresent,
            ShrinkTextToFit = ShrinkTextToFit.IsPresent,
            MinimumShrinkFontSize = MinimumShrinkFontSize,
            KeepTogether = KeepTogether.IsPresent,
            KeepWithNext = KeepWithNext.IsPresent,
            NoBorder = NoBorder.IsPresent,
            NoHeaderFill = NoHeaderFill.IsPresent,
            NoRowStripeFill = NoRowStripeFill.IsPresent,
            HeaderRowCount = HeaderRowCount,
            RepeatHeaderRowCount = RepeatHeaderRowCount,
            FooterRowCount = FooterRowCount
        });
    }

    private static object[] BuildRows(object? value)
    {
        var items = new List<object?>();
        TableInputCollector.AddInput(items, value);
        return TableInputCollector.RequireRows(items, nameof(InputObject));
    }
}
