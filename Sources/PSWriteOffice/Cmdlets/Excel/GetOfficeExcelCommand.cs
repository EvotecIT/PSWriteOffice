using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Opens an existing Excel workbook.</summary>
/// <para>Returns the underlying <see cref="ExcelDocument"/> so callers can inspect or reuse it in DSL pipelines.</para>
/// <example>
///   <summary>Load a workbook in read-only mode.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$workbook = Get-OfficeExcel -Path .\report.xlsx -ReadOnly</code>
///   <para>Loads <c>report.xlsx</c> for inspection without enabling writes.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcel", DefaultParameterSetName = ParameterSetPath)]
public sealed class GetOfficeExcelCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetUri = "Uri";

    /// <summary>Path to the workbook to load.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("Path", "FilePath")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Remote workbook URI to load.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetUri)]
    [Alias("Url")]
    public Uri? Uri { get; set; }

    /// <summary>Allow HTTP workbook downloads in addition to HTTPS.</summary>
    [Parameter(ParameterSetName = ParameterSetUri)]
    public SwitchParameter AllowHttp { get; set; }

    /// <summary>Open the file in read-only mode.</summary>
    [Parameter]
    public SwitchParameter ReadOnly { get; set; }

    /// <summary>Enable automatic saves on the underlying document.</summary>
    [Parameter]
    public SwitchParameter AutoSave { get; set; }

    /// <summary>Password used to open an encrypted workbook package.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (ParameterSetName == ParameterSetUri)
        {
            if (AutoSave.IsPresent)
            {
                throw new PSArgumentException("Remote workbooks cannot be opened with AutoSave. Save to a local path explicitly after loading.");
            }

            if (Uri == null)
            {
                throw new PSArgumentException("Workbook URI was not provided.", nameof(Uri));
            }

            var remoteDocument = ExcelDocumentService.LoadDocument(Uri, ReadOnly.IsPresent, AllowHttp.IsPresent, Password);
            WriteObject(remoteDocument);
            return;
        }

        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
        }

        ExcelDocument document = ExcelDocumentService.LoadDocument(resolvedPath, ReadOnly.IsPresent, AutoSave.IsPresent, Password);
        WriteObject(document);
    }
}
