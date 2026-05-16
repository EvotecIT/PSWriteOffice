using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Models.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets document statistics from a Word document.</summary>
/// <para>Returns a PowerShell-friendly snapshot of OfficeIMO.Word statistics for quick reporting and validation.</para>
[Cmdlet(VerbsCommon.Get, "OfficeWordStatistics", DefaultParameterSetName = ParameterSetPath)]
[Alias("WordStatistics")]
[OutputType(typeof(WordDocumentStatisticsInfo))]
public sealed class GetOfficeWordStatisticsCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the Word document.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath", "Path")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Document to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WordDocument? document = null;
        var dispose = false;

        try
        {
            if (ParameterSetName == ParameterSetPath)
            {
                var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
                if (!File.Exists(resolvedPath))
                {
                    throw new FileNotFoundException($"File {resolvedPath} doesn't exist.", resolvedPath);
                }
                document = WordDocumentService.LoadDocument(resolvedPath, readOnly: true, autoSave: false);
                dispose = true;
            }
            else
            {
                document = Document;
            }

            var statistics = document.Statistics ?? new WordDocumentStatistics(document);

            WriteObject(new WordDocumentStatisticsInfo(statistics));
        }
        finally
        {
            if (dispose)
            {
                document?.Dispose();
            }
        }
    }
}
