using System.Management.Automation;
using OfficeIMO.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Searches normalized document blocks and returns Reader-owned page citations for each match.</summary>
/// <example>
///   <summary>Find text and inspect its page citations.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$document = Get-OfficeDocument -Path .\Policy.docx -IncludePageLocations
/// $matches = $document | Search-OfficeDocument -Query 'retention period'
/// $matches.Hits | Select-Object -ExpandProperty Pages</code>
///   <para>Uses OfficeIMO.Reader search and location contracts without reparsing document text in PowerShell.</para>
/// </example>
[Cmdlet(VerbsCommon.Search, "OfficeDocument")]
[OutputType(typeof(OfficeDocumentSearchResult))]
public sealed class SearchOfficeDocumentCommand : PSCmdlet
{
    /// <summary>Normalized document returned by Get-OfficeDocument.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public OfficeDocumentReadResult InputObject { get; set; } = null!;

    /// <summary>Text to find in normalized document blocks.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [ValidateNotNullOrEmpty]
    public string Query { get; set; } = string.Empty;

    /// <summary>Use case-sensitive ordinal matching.</summary>
    [Parameter]
    public SwitchParameter MatchCase { get; set; }

    /// <summary>Return only occurrences surrounded by non-word characters.</summary>
    [Parameter]
    public SwitchParameter WholeWord { get; set; }

    /// <summary>Maximum number of occurrences to return.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int MaximumResults { get; set; } = 1000;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WriteObject(InputObject.Search(Query, new OfficeDocumentSearchOptions
        {
            MatchCase = MatchCase.IsPresent,
            WholeWord = WholeWord.IsPresent,
            MaximumResults = MaximumResults
        }));
    }
}
