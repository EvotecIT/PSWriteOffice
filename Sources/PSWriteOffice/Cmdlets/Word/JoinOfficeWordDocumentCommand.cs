using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Appends one or more Word documents into a base Word document.</summary>
/// <para>Uses OfficeIMO.Word document append support and preserves the wrapper as an operator-friendly merge command.</para>
[Cmdlet(VerbsCommon.Join, "OfficeWordDocument", DefaultParameterSetName = ParameterSetPath)]
[Alias("Merge-OfficeWordDocument", "WordDocumentJoin")]
[OutputType(typeof(WordDocument))]
public sealed class JoinOfficeWordDocumentCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Base document path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("Path", "BasePath")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Base document object.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Documents to append to the base document.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [Alias("SourcePath")]
    public string[] AppendPath { get; set; } = Array.Empty<string>();

    /// <summary>Optional output path. When omitted for path input, the base document is updated in place.</summary>
    [Parameter]
    public string? OutputPath { get; set; }

    /// <summary>Open the saved output with the shell.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <summary>Emit the merged Word document instead of disposing it.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WordDocument? document = null;
        var dispose = false;

        try
        {
            if (ParameterSetName == ParameterSetPath)
            {
                var resolvedPath = ResolveExistingPath(InputPath);
                document = WordDocumentService.LoadDocument(resolvedPath, readOnly: false, autoSave: false);
                dispose = true;
            }
            else
            {
                document = Document;
            }

            foreach (var sourcePath in AppendPath)
            {
                using var source = WordDocumentService.LoadDocument(ResolveExistingPath(sourcePath), readOnly: true, autoSave: false);
                document.AppendDocument(source);
            }

            string? savedPath = null;

            if (!string.IsNullOrWhiteSpace(OutputPath))
            {
                savedPath = ResolveOutputPath(OutputPath!);
                document.Save(savedPath, false);
            }
            else if (ParameterSetName == ParameterSetPath)
            {
                document.Save(false);
                savedPath = document.FilePath;
            }
            else if (Show.IsPresent)
            {
                savedPath = document.FilePath ?? throw new InvalidOperationException("No saved file path was available.");
                document.Save(false);
            }

            if (Show.IsPresent)
            {
                FileOpenService.Open(savedPath ?? throw new InvalidOperationException("No saved file path was available."));
            }

            if (PassThru.IsPresent)
            {
                dispose = false;
                WriteObject(document);
            }
        }
        finally
        {
            if (dispose)
            {
                document?.Dispose();
            }
        }
    }

    private string ResolveExistingPath(string path)
    {
        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(path);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"File {resolvedPath} doesn't exist.", resolvedPath);
        }
        return resolvedPath;
    }

    private string ResolveOutputPath(string path)
    {
        var providerPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(path);
        return System.IO.Path.IsPathRooted(providerPath)
            ? providerPath
            : System.IO.Path.Combine(SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }
}
