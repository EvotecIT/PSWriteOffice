using System.Management.Automation;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates a reusable Word table cell definition for explicit table rows.</summary>
/// <example>
///   <summary>Create a full-width Word table section row.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$row = @(New-OfficeWordTableCell -Text 'Identity systems' -ColumnSpan 3)</code>
///   <para>The returned cell can be passed to WordTable inside explicit row arrays.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeWordTableCell")]
[OutputType(typeof(OfficeTableCellSpec))]
public sealed class NewOfficeWordTableCellCommand : PSCmdlet
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
