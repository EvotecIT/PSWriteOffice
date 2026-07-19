using System.Collections;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Fills and optionally flattens simple AcroForm fields in an existing PDF.</summary>
/// <example>
///   <summary>Fill and flatten a PDF form.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$fields = @{
///     Requester = 'Ada Lovelace'
///     Priority = 'High'
///     Approved = $true
/// }
/// Set-OfficePdfForm -Path .\Examples\Documents\Request.pdf -OutputPath .\Examples\Documents\Request-FilledFlat.pdf -Field $fields -Flatten</code>
///   <para>Fills simple AcroForm fields and writes a flattened PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfForm", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class SetOfficePdfFormCommand : PSCmdlet
{
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Field values keyed by form field name.</summary>
    [Parameter]
    public Hashtable? Field { get; set; }

    /// <summary>Flatten simple form fields after filling, or flatten without filling when -Field is omitted.</summary>
    [Parameter]
    public SwitchParameter Flatten { get; set; }

    /// <summary>True to keep /NeedAppearances enabled for legacy PDF viewers after filling fields.</summary>
    [Parameter]
    public SwitchParameter KeepNeedAppearances { get; set; }

    /// <summary>Append simple form field values as an incremental PDF revision instead of rewriting the existing PDF.</summary>
    [Parameter]
    public SwitchParameter Incremental { get; set; }

    /// <summary>TrueType or OpenType/CFF font file used to synthesize Unicode form field appearances.</summary>
    [Parameter]
    public string? AppearanceFontPath { get; set; }

    /// <summary>PDF font family name used for the supplied appearance font.</summary>
    [Parameter]
    public string? AppearanceFontFamilyName { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var inputPath = PdfCommandUtilities.ResolvePath(this, Path);
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (Incremental.IsPresent)
        {
            if (Flatten.IsPresent)
            {
                throw new PSArgumentException("-Incremental cannot be combined with -Flatten because flattening requires a full rewrite.");
            }

            if (!string.IsNullOrWhiteSpace(AppearanceFontPath))
            {
                throw new PSArgumentException("-Incremental uses built-in Helvetica appearance streams; use -KeepNeedAppearances or a full rewrite when custom appearance fonts are required.");
            }

            if (Field == null || Field.Count == 0)
            {
                throw new PSArgumentException("Provide -Field values when using -Incremental.", nameof(Field));
            }

            if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write incrementally updated PDF form"))
            {
                return;
            }

            PdfCommandUtilities.EnsureDirectory(outputPath);
            var options = new PdfIncrementalFormFieldUpdateOptions
            {
                KeepNeedAppearances = KeepNeedAppearances.IsPresent,
                GenerateAppearanceStreams = !KeepNeedAppearances.IsPresent
            };
            PdfDocument
                .Open(inputPath)
                .Forms.AppendRevision(PdfCommandUtilities.ConvertFieldValues(Field), options)
                .Save(outputPath)
                .RequireSuccess();
            WriteObject(new FileInfo(outputPath));
            return;
        }

        var document = PdfDocument.Open(inputPath);
        var formOptions = PdfCommandUtilities.CreateFormFillerOptions(this, AppearanceFontPath, AppearanceFontFamilyName, KeepNeedAppearances.IsPresent);
        PdfDocument result;
        if (Field == null || Field.Count == 0)
        {
            if (!Flatten.IsPresent)
            {
                throw new PSArgumentException("Provide -Field values or use -Flatten.", nameof(Field));
            }

            result = formOptions == null
                ? document.Forms.Flatten()
                : document.Forms.Flatten(formOptions);
        }
        else
        {
            var values = PdfCommandUtilities.ConvertFieldValues(Field);
            if (Flatten.IsPresent)
            {
                result = formOptions == null
                    ? document.Forms.FillAndFlatten(values)
                    : document.Forms.FillAndFlatten(values, formOptions);
            }
            else
            {
                result = formOptions == null
                    ? document.Forms.Fill(values)
                    : document.Forms.Fill(values, formOptions);
            }
        }

        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write updated PDF form"))
        {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        result.Save(outputPath).RequireSuccess();
        WriteObject(new FileInfo(outputPath));
    }
}
