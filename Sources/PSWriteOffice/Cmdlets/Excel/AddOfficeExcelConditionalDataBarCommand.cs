using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a data bar conditional format to a range.</summary>
/// <example>
///   <summary>Add blue data bars.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelConditionalDataBar -Range 'D2:D50' -Color '#4F81BD' }</code>
///   <para>Applies data bars to column D.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelConditionalDataBar", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelConditionalDataBar")]
public sealed class AddOfficeExcelConditionalDataBarCommand : PSCmdlet
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

    /// <summary>Bar color in hex (#RRGGBB or FFRRGGBB).</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Color { get; set; } = string.Empty;

    /// <summary>Emit the range after applying the format.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();

        sheet.AddConditionalDataBar(Range, NormalizeColor(Color));

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
