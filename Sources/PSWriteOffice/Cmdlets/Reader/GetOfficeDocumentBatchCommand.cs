using System.Collections.Generic;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Reads a bounded set of documents asynchronously while retaining input order.</summary>
/// <example>
///   <summary>Read documents with four operations in flight.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$batch = [OfficeIMO.Reader.ReaderBatchOptions]::new(); $batch.MaxDegreeOfParallelism = 4; Get-ChildItem .\Reports -File | Get-OfficeDocumentBatch -BatchOptions $batch</code>
///   <para>OfficeIMO.Reader bounds concurrency and returns results in pipeline input order.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeDocumentBatch")]
[OutputType(typeof(OfficeDocumentReadResult))]
public sealed class GetOfficeDocumentBatchCommand : AsyncPSCmdlet
{
    private readonly List<string> _paths = new();

    /// <summary>Paths to read.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
    [Alias("FullName", "FilePath")]
    public string[] Path { get; set; } = System.Array.Empty<string>();

    /// <summary>Optional immutable reader with caller-configured processors.</summary>
    [Parameter]
    public OfficeDocumentReader? Reader { get; set; }

    /// <summary>Optional source-reading limits and format behavior.</summary>
    [Parameter]
    public ReaderOptions? ReaderOptions { get; set; }

    /// <summary>Optional maximum document count and degree of parallelism.</summary>
    [Parameter]
    public ReaderBatchOptions? BatchOptions { get; set; }

    /// <inheritdoc />
    protected override Task ProcessRecordAsync()
    {
        foreach (var path in Path)
        {
            _paths.Add(ReaderCommandUtilities.ResolvePath(this, path));
        }
        return Task.CompletedTask;
    }

    /// <inheritdoc />
    protected override async Task EndProcessingAsync()
    {
        var reader = Reader ?? ReaderCommandUtilities.Reader;
        var results = await reader.ReadDocumentsAsync(_paths, ReaderOptions, BatchOptions, CancelToken).ConfigureAwait(false);
        foreach (var result in results) WriteObject(result);
    }
}
