using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets prompt and error messages on existing Excel data validation rules.</summary>
/// <example>
///   <summary>Add prompt and error text to a validation rule.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rules = Set-OfficeExcelDataValidationMessage -Path .\Report.xlsx -Sheet Data -HeaderName Sales -TableName ServiceHealth -PromptTitle 'Sales' -Prompt 'Enter 1-1000' -ErrorTitle 'Invalid sales' -ErrorMessage 'Enter a whole number from 1 to 1000' -ShowInputMessage -ShowErrorMessage -PassThru
/// $rules |
///     Select-Object Range, PromptTitle, ErrorTitle</code>
///   <para>Updates validation metadata for matching rules and saves the workbook.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelDataValidationMessage", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Low, DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelDataValidationMessage")]
[OutputType(typeof(PSObject))]
public sealed class SetOfficeExcelDataValidationMessageCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("Path", "FilePath")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name to update. Defaults to the current DSL sheet.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) to update. Defaults to the current DSL sheet.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>A1 range used to select existing validation rules.</summary>
    [Parameter]
    public string? Range { get; set; }

    /// <summary>Header or table column name used to resolve the validation rules to update.</summary>
    [Parameter]
    [Alias("ColumnName")]
    public string? HeaderName { get; set; }

    /// <summary>Optional table name for header-based range resolution.</summary>
    [Parameter]
    public string? TableName { get; set; }

    /// <summary>Worksheet header row used when resolving HeaderName without a table. Use 0 for the first row of the used range.</summary>
    [Parameter]
    public int HeaderRow { get; set; }

    /// <summary>Include the header cell in the resolved range.</summary>
    [Parameter]
    public SwitchParameter IncludeHeader { get; set; }

    /// <summary>Input prompt title. Omit or pass null to clear the title.</summary>
    [Parameter]
    public string? PromptTitle { get; set; }

    /// <summary>Input prompt text. Omit or pass null to clear the prompt.</summary>
    [Parameter]
    public string? Prompt { get; set; }

    /// <summary>Error title. Omit or pass null to clear the title.</summary>
    [Parameter]
    public string? ErrorTitle { get; set; }

    /// <summary>Error message text. Omit or pass null to clear the message.</summary>
    [Parameter]
    [Alias("Error")]
    public string? ErrorMessage { get; set; }

    /// <summary>Forces Excel to show the input prompt.</summary>
    [Parameter]
    public SwitchParameter ShowInputMessage { get; set; }

    /// <summary>Forces Excel to show the validation error.</summary>
    [Parameter]
    public SwitchParameter ShowErrorMessage { get; set; }

    /// <summary>Returns matching validation rules after updating them.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (!HasMessageOption())
        {
            throw new PSArgumentException("Specify at least one prompt, error, or display option.");
        }

        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, InputPath, Document, readOnly: false);
        var sheet = ExcelWorkbookCommandService.ResolveSheet(this, workbook.Document, ParameterSetName, Sheet, SheetIndex);
        string targetRange = ExcelTargetRangeResolver.Resolve(sheet, Range, HeaderName, TableName, HeaderRow, IncludeHeader.IsPresent);
        var target = $"{sheet.Name}!{targetRange}";
        if (!ShouldProcess(target, "Set Excel data validation messages"))
        {
            return;
        }

        var existing = sheet.GetDataValidations(targetRange).FirstOrDefault();
        var promptTitle = ResolveMessageValue(nameof(PromptTitle), PromptTitle, existing?.PromptTitle);
        var prompt = ResolveMessageValue(nameof(Prompt), Prompt, existing?.Prompt);
        var errorTitle = ResolveMessageValue(nameof(ErrorTitle), ErrorTitle, existing?.ErrorTitle);
        var errorMessage = ResolveMessageValue(nameof(ErrorMessage), ErrorMessage, existing?.Error);
        var displayState = ExcelDataValidationMessageDisplayState.Capture(sheet, targetRange);
        bool? boundShowInputMessage = ResolveBoundDisplayFlag(nameof(ShowInputMessage), ShowInputMessage);
        bool? boundShowErrorMessage = ResolveBoundDisplayFlag(nameof(ShowErrorMessage), ShowErrorMessage);
        bool showInputMessage = boundShowInputMessage ?? ResolveUnboundDisplayFlag(displayState.FirstShowInputMessage, displayState.FirstHasInputMessageText, promptTitle, prompt);
        bool showErrorMessage = boundShowErrorMessage ?? ResolveUnboundDisplayFlag(displayState.FirstShowErrorMessage, displayState.FirstHasErrorMessageText, errorTitle, errorMessage);
        sheet.SetDataValidationMessages(targetRange, new ExcelDataValidationMessageOptions
        {
            PromptTitle = promptTitle,
            Prompt = prompt,
            ErrorTitle = errorTitle,
            Error = errorMessage,
            ShowInputMessage = showInputMessage,
            ShowErrorMessage = showErrorMessage
        });
        displayState.Restore(sheet, targetRange, boundShowInputMessage, boundShowErrorMessage);

        workbook.SaveIfOwned();

        if (PassThru.IsPresent)
        {
            var path = string.Equals(ParameterSetName, ParameterSetPath, StringComparison.OrdinalIgnoreCase)
                ? InputPath
                : null;
            foreach (var validation in sheet.GetDataValidations(targetRange).Select(validation => ExcelRuleRecordService.CreateDataValidationRecord(validation, sheet.Name, path)))
            {
                WriteObject(validation);
            }
        }
    }

    private bool HasMessageOption()
    {
        return MyInvocation.BoundParameters.ContainsKey(nameof(PromptTitle))
            || MyInvocation.BoundParameters.ContainsKey(nameof(Prompt))
            || MyInvocation.BoundParameters.ContainsKey(nameof(ErrorTitle))
            || MyInvocation.BoundParameters.ContainsKey(nameof(ErrorMessage))
            || MyInvocation.BoundParameters.ContainsKey(nameof(ShowInputMessage))
            || MyInvocation.BoundParameters.ContainsKey(nameof(ShowErrorMessage));
    }

    private string? ResolveMessageValue(string parameterName, string? value, string? existing)
    {
        return MyInvocation.BoundParameters.ContainsKey(parameterName) ? value : existing;
    }

    private bool? ResolveBoundDisplayFlag(string parameterName, SwitchParameter value)
    {
        return MyInvocation.BoundParameters.ContainsKey(parameterName)
            ? value.IsPresent
            : null;
    }

    private static bool ResolveUnboundDisplayFlag(bool? existingDisplayFlag, bool existingHadMessageText, string? title, string? message)
    {
        if (existingHadMessageText)
        {
            return existingDisplayFlag ?? false;
        }

        return !string.IsNullOrEmpty(title) || !string.IsNullOrEmpty(message);
    }
}
