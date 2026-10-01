using System;
using System.Management.Automation;
using System.Threading;
using System.Threading.Tasks;
#if !FRAMEWORK
using OfficeIMO.Workflows;
#endif
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Converts a local DOC/DOCX/TXT directory to PDF with durable restart checks. Requires PowerShell 7.</summary>
/// <para>Source, output and checkpoint directories must be separate. Rerunning verifies completed source and output hashes.
/// Each output retains its full source filename plus .pdf. Known DOC import loss blocks output unless explicitly accepted.</para>
/// <example><summary>Run or resume an archive.</summary><prefix>PS&gt; </prefix>
/// <code>Export-OfficePdfArchive -InputDirectory .\Documents -OutputDirectory .\PDF -CheckpointDirectory .\PDF-State</code></example>
[Cmdlet(VerbsData.Export, "OfficePdfArchive", SupportsShouldProcess = true)]
#if !FRAMEWORK
[OutputType(typeof(OfficePdfArchiveResult))]
#endif
public sealed class ExportOfficePdfArchiveCommand : AsyncPSCmdlet {
    /// <summary>Local DOC, DOCX and TXT source tree.</summary>
    [Parameter(Mandatory = true)] public string InputDirectory { get; set; } = string.Empty;
    /// <summary>Separate PDF destination tree.</summary>
    [Parameter(Mandatory = true)] public string OutputDirectory { get; set; } = string.Empty;
    /// <summary>Separate durable checkpoint tree.</summary>
    [Parameter(Mandatory = true)] public string CheckpointDirectory { get; set; } = string.Empty;
    /// <summary>Maximum simultaneously executing files.</summary>
    [Parameter] [ValidateRange(1, 8)] public int MaximumConcurrency { get; set; } = 2;
    /// <summary>Maximum selected source files.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaximumFiles { get; set; } = 1_000_000;
    /// <summary>Maximum source bytes per file.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long MaximumInputBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum PDF bytes per file.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long MaximumOutputBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Explicit literal text encoding.</summary>
    [Parameter] public string? TextEncoding { get; set; }
    /// <summary>Literal text tab columns.</summary>
    [Parameter] [ValidateRange(1, 32)] public int TabSize { get; set; } = 8;
    /// <summary>Maximum decoded and tab-expanded characters per TXT source.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaximumTextCharacters { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum generated pages per TXT source.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaximumTextPages { get; set; } = 10_000;
    /// <summary>Retry recorded failures, including corrected failed sources. Completed items remain protected.</summary>
    [Parameter] public SwitchParameter RetryFailed { get; set; }
    /// <summary>Accept reported legacy DOC import loss.</summary>
    [Parameter] public SwitchParameter AllowLegacyImportLoss { get; set; }
    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
#if FRAMEWORK
        await Task.CompletedTask;
        throw new PlatformNotSupportedException("PDF archive processing requires PowerShell 7.4 or newer. Single-file export is available through Export-OfficeDocumentPdf.");
#else
        var request = new OfficePdfArchiveRequest {
            InputDirectory = PdfCommandUtilities.ResolvePath(this, InputDirectory),
            OutputDirectory = PdfCommandUtilities.ResolvePath(this, OutputDirectory),
            CheckpointDirectory = PdfCommandUtilities.ResolvePath(this, CheckpointDirectory),
            MaximumConcurrency = MaximumConcurrency, MaximumFiles = MaximumFiles,
            MaximumInputBytes = MaximumInputBytes, MaximumOutputBytes = MaximumOutputBytes,
            TextEncoding = TextEncoding, TabSize = TabSize, RetryFailed = RetryFailed.IsPresent,
            MaximumTextCharacters = MaximumTextCharacters, MaximumTextPages = MaximumTextPages,
            AllowLegacyImportLoss = AllowLegacyImportLoss.IsPresent
        };
        if (!ShouldProcess(request.OutputDirectory, "Convert or resume DOC/DOCX/TXT archive; write PDF and checkpoint files")) return;
        var progress = new FailureProgress(this);
        OfficePdfArchiveResult result = await OfficePdfArchiveWorkflow.RunAsync(request, progress, CancelToken);
        if (progress.Failures > 100) WriteWarning($"{progress.Failures - 100} further failure messages were suppressed. Check the returned failure count and per-file checkpoints.");
        WriteObject(result);
#endif
    }
#if !FRAMEWORK
    private sealed class FailureProgress(ExportOfficePdfArchiveCommand command) : IProgress<OfficePdfArchiveItemResult> {
        private int _failures;
        public int Failures => Volatile.Read(ref _failures);
        public void Report(OfficePdfArchiveItemResult item) {
            if (item.Status == OfficeWorkflowStatus.Failed && Interlocked.Increment(ref _failures) <= 100)
                command.WriteWarning(item.InputPath + ": " + item.Summary);
        }
    }
#endif
}
