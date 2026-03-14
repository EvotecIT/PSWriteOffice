using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a two-color scale conditional format to a range.</summary>
/// <example>
///   <summary>Apply a red-to-green color scale.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelConditionalColorScale -Range 'B2:B50' -StartColor '#FF0000' -EndColor '#00FF00' }</code>
///   <para>Applies a red-to-green scale to column B.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelConditionalColorScale", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelConditionalColorScale")]
public sealed class AddOfficeExcelConditionalColorScaleCommand : PSCmdlet
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

    /// <summary>Start color in hex (#RRGGBB or FFRRGGBB).</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string StartColor { get; set; } = string.Empty;

    /// <summary>End color in hex (#RRGGBB or FFRRGGBB).</summary>
    [Parameter(Mandatory = true, Position = 2)]
    public string EndColor { get; set; } = string.Empty;

    /// <summary>Emit the range after applying the format.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();

        sheet.AddConditionalColorScale(Range, NormalizeColor(StartColor), NormalizeColor(EndColor));

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

    private static string NormalizeColor(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new PSArgumentException("Color cannot be empty.");
        }

        var trimmed = value.Trim().TrimStart('#');
        if (trimmed.Length == 6)
        {
            return "FF" + trimmed.ToUpperInvariant();
        }

        if (trimmed.Length == 8)
        {
            return trimmed.ToUpperInvariant();
        }

        throw new PSArgumentException("Color must be in #RRGGBB or FFRRGGBB format.");
    }
}
