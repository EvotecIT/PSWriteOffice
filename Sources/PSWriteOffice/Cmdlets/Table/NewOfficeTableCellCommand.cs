using System.Management.Automation;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Table;

/// <summary>Creates a reusable table cell definition for Word and PDF table cmdlets.</summary>
/// <example>
///   <summary>Create a full-width section row.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$row = @(New-OfficeTableCell -Text 'Identity systems' -ColumnSpan 3)</code>
///   <para>The returned cell can be passed to WordTable or PdfTable inside explicit row arrays.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeTableCell")]
[Alias("OfficeTableCell")]
[OutputType(typeof(OfficeTableCellSpec))]
public sealed class NewOfficeTableCellCommand : PSCmdlet
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
