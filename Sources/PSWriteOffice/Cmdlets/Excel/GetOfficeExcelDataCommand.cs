using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Reads worksheet data as dictionaries or PSCustomObjects.</summary>
/// <para>Uses the first row as headers and materializes rows via the OfficeIMO Excel fluent reader.</para>
/// <example>
///   <summary>Read the used range as PSCustomObjects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeExcelData -Path .\report.xlsx -Sheet 'Summary' | Format-Table</code>
///   <para>Returns each row as a PSCustomObject with properties mapped from the header row.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelData")]
public sealed class GetOfficeExcelDataCommand : PSCmdlet
{
    /// <summary>Workbook to read when it is already open.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Path to the workbook when no <see cref="Document"/> is supplied.</summary>
    [Parameter(Position = 0)]
    public string? Path { get; set; }

    /// <summary>Worksheet name to read; defaults to the first sheet.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Optional A1 range (e.g. A1:D10). When omitted, the sheet's used range is read.</summary>
    [Parameter]
    public string? Range { get; set; }

    /// <summary>Prefer decimals (instead of doubles) for numeric cells.</summary>
    [Parameter]
    public SwitchParameter NumericAsDecimal { get; set; }

    /// <summary>Emit each row as a dictionary instead of PSCustomObjects.</summary>
    [Parameter]
    public SwitchParameter AsHashtable { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = Document;
        var ownsDocument = false;

        if (document == null)
        {
            if (string.IsNullOrWhiteSpace(Path))
            {
                throw new PSArgumentException("Specify -Path or provide an ExcelDocument on the pipeline.");
            }

            var resolved = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            if (!File.Exists(resolved))
            {
                throw new FileNotFoundException($"File '{resolved}' was not found.", resolved);
            }

            document = ExcelDocument.Load(resolved, readOnly: true);
            ownsDocument = true;
        }

        try
        {
            var fluent = document.Read();
            ExcelFluentReadSheet sheetScope = string.IsNullOrWhiteSpace(Sheet)
                ? fluent.Sheet(0)
                : fluent.Sheet(Sheet!);

            ExcelFluentReadRange rangeScope = string.IsNullOrWhiteSpace(Range)
                ? sheetScope.UsedRange()
                : sheetScope.Range(Range!);

            if (NumericAsDecimal.IsPresent)
            {
                rangeScope.NumericAsDecimal();
            }

            foreach (var row in rangeScope.AsRows())
            {
                if (row == null)
                {
                    continue;
                }

                if (AsHashtable.IsPresent)
                {
                    WriteObject(row);
                }
                else
                {
                    var psObj = new PSObject();
                    foreach (KeyValuePair<string, object?> kv in row)
                    {
                        psObj.Properties.Add(new PSNoteProperty(kv.Key, kv.Value));
                    }
                    WriteObject(psObj);
                }
            }
        }
        finally
        {
            if (ownsDocument)
            {
                document.Dispose();
            }
        }
    }
}
