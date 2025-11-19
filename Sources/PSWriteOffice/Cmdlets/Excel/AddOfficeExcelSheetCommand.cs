using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds or reuses a worksheet within the current Excel DSL scope.</summary>
/// <para>Creates the sheet when missing, pushes it onto the DSL stack, and executes the nested script block.</para>
/// <example>
///   <summary>Create a sheet named Data.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeExcel -Path .\report.xlsx { Add-OfficeExcelSheet -Name 'Data' { ExcelCell -Address 'A1' -Value 'Region' } }</code>
///   <para>Creates a workbook with a worksheet named Data and writes the header “Region”.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelSheet")]
[Alias("ExcelSheet")]
public sealed class AddOfficeExcelSheetCommand : PSCmdlet
{
    /// <summary>Name of the worksheet to create or reuse. When omitted the last sheet is reused or a default sheet is created.</summary>
    [Parameter(Position = 0)]
    public string? Name { get; set; }

    /// <summary>Controls how invalid sheet names are handled.</summary>
    [Parameter]
    public SheetNameValidationMode ValidationMode { get; set; } = SheetNameValidationMode.Sanitize;

    /// <summary>Code to execute inside the worksheet context.</summary>
    [Parameter(Position = 1, ValueFromRemainingArguments = true)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Emit the <see cref="ExcelSheet"/> object after execution.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = ExcelDslContext.Require(this);
        var sheet = context.Document.GetOrCreateSheet(Name, ValidationMode);

        using (context.Push(sheet))
        {
            Content?.InvokeReturnAsIs();
        }

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }
}
