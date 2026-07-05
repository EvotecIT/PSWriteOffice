using System;
using System.Collections.Concurrent;
using System.Management.Automation;
using System.Threading;
using System.Threading.Tasks;

namespace PSWriteOffice;
/// <summary>
/// Base class for cmdlets that await asynchronous engine work while routing PowerShell pipeline writes
/// back through the synchronous cmdlet pipeline thread.
/// </summary>
public abstract class AsyncPSCmdlet : PSCmdlet, IDisposable {
    private enum PipelineType {
        Output,
        OutputEnumerate,
        Error,
        Warning,
        Verbose,
        Debug,
        Information,
        Progress,
        ShouldProcess,
        PromptForCredential
    }

    private readonly CancellationTokenSource _cancelSource = new();
    private BlockingCollection<(object? Value, PipelineType Type)>? _currentOutPipe;
    private BlockingCollection<object?>? _currentReplyPipe;
    private int _pipelineThreadId;

    /// <summary>Cancellation token triggered when PowerShell stops the cmdlet.</summary>
    protected internal CancellationToken CancelToken => _cancelSource.Token;

    /// <inheritdoc />
    protected override void BeginProcessing()
        => RunBlockInAsync(BeginProcessingAsync);

    /// <summary>Asynchronous begin hook.</summary>
    protected virtual Task BeginProcessingAsync()
        => Task.CompletedTask;

    /// <inheritdoc />
    protected override void ProcessRecord()
        => RunBlockInAsync(ProcessRecordAsync);

    /// <summary>Asynchronous process-record hook.</summary>
    protected virtual Task ProcessRecordAsync()
        => Task.CompletedTask;

    /// <inheritdoc />
    protected override void EndProcessing()
        => RunBlockInAsync(EndProcessingAsync);

    /// <summary>Asynchronous end hook.</summary>
    protected virtual Task EndProcessingAsync()
        => Task.CompletedTask;

    /// <inheritdoc />
    protected override void StopProcessing()
        => _cancelSource.Cancel();

    /// <summary>Thread-safe ShouldProcess bridge for asynchronous cmdlet code.</summary>
    public new bool ShouldProcess(string? target, string action) {
        ThrowIfStopped();
        if (_currentOutPipe is null || _currentReplyPipe is null || IsPipelineThread) {
            return base.ShouldProcess(target ?? string.Empty, action);
        }

        _currentOutPipe.Add(((target ?? string.Empty, action), PipelineType.ShouldProcess), CancelToken);
        return (bool)_currentReplyPipe.Take(CancelToken)!;
    }

    /// <summary>Thread-safe credential prompt bridge for asynchronous cmdlet code.</summary>
    public PSCredential? PromptForCredential(string caption, string message, string userName, string targetName) {
        ThrowIfStopped();
        if (_currentOutPipe is null || _currentReplyPipe is null || IsPipelineThread) {
            return Host.UI.PromptForCredential(caption, message, userName, targetName);
        }

        _currentOutPipe.Add(((caption, message, userName, targetName), PipelineType.PromptForCredential), CancelToken);
        return (PSCredential?)_currentReplyPipe.Take(CancelToken);
    }

    /// <summary>Thread-safe output bridge for asynchronous cmdlet code.</summary>
    public new void WriteObject(object? sendToPipeline)
        => WriteObject(sendToPipeline, false);

    /// <summary>Thread-safe output bridge for asynchronous cmdlet code.</summary>
    public new void WriteObject(object? sendToPipeline, bool enumerateCollection) {
        ThrowIfStopped();
        if (_currentOutPipe is null || IsPipelineThread) {
            base.WriteObject(sendToPipeline, enumerateCollection);
            return;
        }

        _currentOutPipe.Add((sendToPipeline, enumerateCollection ? PipelineType.OutputEnumerate : PipelineType.Output), CancelToken);
    }

    /// <summary>Thread-safe error bridge for asynchronous cmdlet code.</summary>
    public new void WriteError(ErrorRecord errorRecord) {
        ThrowIfStopped();
        if (_currentOutPipe is null || IsPipelineThread) {
            base.WriteError(errorRecord);
            return;
        }

        _currentOutPipe.Add((errorRecord, PipelineType.Error), CancelToken);
    }

    /// <summary>Thread-safe warning bridge for asynchronous cmdlet code.</summary>
    public new void WriteWarning(string message) {
        ThrowIfStopped();
        if (_currentOutPipe is null || IsPipelineThread) {
            base.WriteWarning(message);
            return;
        }

        _currentOutPipe.Add((message, PipelineType.Warning), CancelToken);
    }

    /// <summary>Thread-safe verbose bridge for asynchronous cmdlet code.</summary>
    public new void WriteVerbose(string message) {
        ThrowIfStopped();
        if (_currentOutPipe is null || IsPipelineThread) {
            base.WriteVerbose(message);
            return;
        }

        _currentOutPipe.Add((message, PipelineType.Verbose), CancelToken);
    }

    /// <summary>Thread-safe debug bridge for asynchronous cmdlet code.</summary>
    public new void WriteDebug(string message) {
        ThrowIfStopped();
        if (_currentOutPipe is null || IsPipelineThread) {
            base.WriteDebug(message);
            return;
        }

        _currentOutPipe.Add((message, PipelineType.Debug), CancelToken);
    }

    /// <summary>Thread-safe information bridge for asynchronous cmdlet code.</summary>
    public new void WriteInformation(InformationRecord informationRecord) {
        ThrowIfStopped();
        if (_currentOutPipe is null || IsPipelineThread) {
            base.WriteInformation(informationRecord);
            return;
        }

        _currentOutPipe.Add((informationRecord, PipelineType.Information), CancelToken);
    }

    /// <summary>Thread-safe progress bridge for asynchronous cmdlet code.</summary>
    public new void WriteProgress(ProgressRecord progressRecord) {
        ThrowIfStopped();
        if (_currentOutPipe is null || IsPipelineThread) {
            base.WriteProgress(progressRecord);
            return;
        }

        _currentOutPipe.Add((progressRecord, PipelineType.Progress), CancelToken);
    }

    /// <summary>Throws when PowerShell has requested cancellation.</summary>
    internal void ThrowIfStopped() {
        if (_cancelSource.IsCancellationRequested) {
            throw new PipelineStoppedException();
        }
    }

    /// <summary>
    /// Disposes managed resources.
    /// </summary>
    public void Dispose() {
        _cancelSource.Dispose();
    }

    private bool IsPipelineThread
        => _pipelineThreadId != 0 && Environment.CurrentManagedThreadId == _pipelineThreadId;

    private void RunBlockInAsync(Func<Task> task) {
        using var outPipe = new BlockingCollection<(object? Value, PipelineType Type)>();
        using var replyPipe = new BlockingCollection<object?>();
        Task blockTask;

        void ClearPipes() {
            _currentOutPipe = null;
            _currentReplyPipe = null;
            _pipelineThreadId = 0;
            CompleteAddingIfNeeded(outPipe);
            CompleteAddingIfNeeded(replyPipe);
        }

        static void CompleteAddingIfNeeded<T>(BlockingCollection<T> pipe) {
            if (!pipe.IsAddingCompleted) {
                pipe.CompleteAdding();
            }
        }

        void PumpItem((object? Value, PipelineType Type) item) {
            switch (item.Type) {
                case PipelineType.Output:
                    base.WriteObject(item.Value);
                    break;
                case PipelineType.OutputEnumerate:
                    base.WriteObject(item.Value, true);
                    break;
                case PipelineType.Error:
                    base.WriteError((ErrorRecord)item.Value!);
                    break;
                case PipelineType.Warning:
                    base.WriteWarning((string)item.Value!);
                    break;
                case PipelineType.Verbose:
                    base.WriteVerbose((string)item.Value!);
                    break;
                case PipelineType.Debug:
                    base.WriteDebug((string)item.Value!);
                    break;
                case PipelineType.Information:
                    base.WriteInformation((InformationRecord)item.Value!);
                    break;
                case PipelineType.Progress:
                    base.WriteProgress((ProgressRecord)item.Value!);
                    break;
                case PipelineType.ShouldProcess:
                    var should = ((string Target, string Action))item.Value!;
                    replyPipe.Add(base.ShouldProcess(should.Target, should.Action), CancelToken);
                    break;
                case PipelineType.PromptForCredential:
                    var prompt = ((string Caption, string Message, string UserName, string TargetName))item.Value!;
                    replyPipe.Add(
                        Host.UI.PromptForCredential(prompt.Caption, prompt.Message, prompt.UserName, prompt.TargetName),
                        CancelToken);
                    break;
            }
        }

        void PumpQueuedItems() {
            while (outPipe.TryTake(out var item)) {
                PumpItem(item);
            }
        }

        _pipelineThreadId = Environment.CurrentManagedThreadId;
        _currentOutPipe = outPipe;
        _currentReplyPipe = replyPipe;

        try {
            blockTask = task();
        } catch {
            ClearPipes();
            throw;
        }

        if (blockTask.IsCompleted) {
            CompleteAddingIfNeeded(outPipe);
            PumpQueuedItems();
            ClearPipes();
            blockTask.GetAwaiter().GetResult();
            return;
        }

        _ = blockTask.ContinueWith(
            _ => ClearPipes(),
            CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously,
            TaskScheduler.Default);

        try {
            foreach (var item in outPipe.GetConsumingEnumerable(CancelToken)) {
                PumpItem(item);
            }
        } catch {
            _cancelSource.Cancel();
            CompleteAddingIfNeeded(outPipe);
            try {
                blockTask.GetAwaiter().GetResult();
            } catch (Exception ex) when (ex is OperationCanceledException or PipelineStoppedException) {
            }

            throw;
        }

        blockTask.GetAwaiter().GetResult();
    }
}
