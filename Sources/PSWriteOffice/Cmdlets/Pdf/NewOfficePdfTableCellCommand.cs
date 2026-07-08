using System.Management.Automation;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates a reusable PDF table cell definition for explicit table rows.</summary>
/// <example>
///   <summary>Create a full-width PDF table section row.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$row = @(New-OfficePdfTableCell -Text 'Identity systems' -ColumnSpan 3)</code>
///   <para>The returned cell can be passed to PdfTable inside explicit row arrays.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePdfTableCell")]
[OutputType(typeof(OfficeTableCellSpec))]
public sealed class NewOfficePdfTableCellCommand : PSCmdlet
{
    /// <summary>Cell text.</summary>
    [Parameter(Position = 0)]
    public string? Text { get; set; }

    /// <summary>Number of logical columns covered by the cell.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int ColumnSpan { get; set; } = 1;

    /// <summary>Number of logical rows covered by the cell.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int RowSpan { get; set; } = 1;

    /// <inheritdoc />
    protected override void ProcessRecord()
        => WriteObject(new OfficeTableCellSpec(Text, ColumnSpan, RowSpan));
}
