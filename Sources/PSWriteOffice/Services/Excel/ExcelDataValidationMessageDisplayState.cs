using System;
using System.Collections.Generic;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal sealed class ExcelDataValidationMessageDisplayState
{
    // OfficeIMO 0.6.48 updates message text through a public API, but it does not expose
    // the existing showInputMessage/showErrorMessage flags in its validation snapshot.
    private static readonly PropertyInfo? WorksheetRootProperty = typeof(ExcelSheet).GetProperty("WorksheetRoot", BindingFlags.Instance | BindingFlags.NonPublic);

    private readonly List<DataValidationDisplaySnapshot> _snapshots;

    private ExcelDataValidationMessageDisplayState(List<DataValidationDisplaySnapshot> snapshots)
    {
        _snapshots = snapshots;
    }

    public bool? FirstShowInputMessage => _snapshots.Count == 0 ? null : _snapshots[0].ShowInputMessage ?? false;

    public bool? FirstShowErrorMessage => _snapshots.Count == 0 ? null : _snapshots[0].ShowErrorMessage ?? false;

    public bool FirstHasInputMessageText => _snapshots.Count > 0 && _snapshots[0].HasInputMessageText;

    public bool FirstHasErrorMessageText => _snapshots.Count > 0 && _snapshots[0].HasErrorMessageText;

    public static ExcelDataValidationMessageDisplayState Capture(ExcelSheet sheet, string targetRange)
    {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (string.IsNullOrWhiteSpace(targetRange)) throw new ArgumentNullException(nameof(targetRange));

        Worksheet worksheet = GetWorksheetRoot(sheet);
        var filter = ParseReferenceArgument(targetRange);
        var snapshots = new List<DataValidationDisplaySnapshot>();
        var validations = worksheet.GetFirstChild<DataValidations>();
        if (validations == null)
        {
            return new ExcelDataValidationMessageDisplayState(snapshots);
        }

        foreach (DataValidation validation in validations.Elements<DataValidation>())
        {
            string range = validation.SequenceOfReferences?.InnerText ?? string.Empty;
            if (!ReferenceListOverlaps(range, filter)) continue;

            snapshots.Add(new DataValidationDisplaySnapshot(
                range,
                validation.ShowInputMessage?.Value,
                validation.ShowErrorMessage?.Value,
                !string.IsNullOrEmpty(validation.PromptTitle?.Value) || !string.IsNullOrEmpty(validation.Prompt?.Value),
                !string.IsNullOrEmpty(validation.ErrorTitle?.Value) || !string.IsNullOrEmpty(validation.Error?.Value)));
        }

        return new ExcelDataValidationMessageDisplayState(snapshots);
    }

    public void Restore(ExcelSheet sheet, string targetRange, bool? showInputMessage, bool? showErrorMessage)
    {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (string.IsNullOrWhiteSpace(targetRange)) throw new ArgumentNullException(nameof(targetRange));

        Worksheet worksheet = GetWorksheetRoot(sheet);
        var filter = ParseReferenceArgument(targetRange);
        var validations = worksheet.GetFirstChild<DataValidations>();
        if (validations == null)
        {
            return;
        }

        var usedSnapshots = new bool[_snapshots.Count];
        bool changed = false;
        foreach (DataValidation validation in validations.Elements<DataValidation>())
        {
            string range = validation.SequenceOfReferences?.InnerText ?? string.Empty;
            if (!ReferenceListOverlaps(range, filter)) continue;

            int snapshotIndex = FindSnapshot(range, usedSnapshots);
            if (snapshotIndex >= 0)
            {
                usedSnapshots[snapshotIndex] = true;
            }

            DataValidationDisplaySnapshot? snapshot = snapshotIndex >= 0 ? _snapshots[snapshotIndex] : null;
            if (showInputMessage.HasValue)
            {
                changed |= SetBoolean(validation, isInputMessage: true, showInputMessage);
            }
            else if (snapshot?.HasInputMessageText == true)
            {
                changed |= SetBoolean(validation, isInputMessage: true, snapshot.ShowInputMessage);
            }

            if (showErrorMessage.HasValue)
            {
                changed |= SetBoolean(validation, isInputMessage: false, showErrorMessage);
            }
            else if (snapshot?.HasErrorMessageText == true)
            {
                changed |= SetBoolean(validation, isInputMessage: false, snapshot.ShowErrorMessage);
            }
        }

        if (changed)
        {
            worksheet.Save();
        }
    }

    private static Worksheet GetWorksheetRoot(ExcelSheet sheet)
    {
        if (WorksheetRootProperty?.GetValue(sheet) is Worksheet worksheet)
        {
            return worksheet;
        }

        throw new InvalidOperationException("Unable to access the worksheet XML for Excel data validation message display state.");
    }

    private int FindSnapshot(string range, bool[] usedSnapshots)
    {
        for (int i = 0; i < _snapshots.Count; i++)
        {
            if (!usedSnapshots[i] && string.Equals(_snapshots[i].Range, range, StringComparison.Ordinal))
            {
                return i;
            }
        }

        return -1;
    }

    private static bool SetBoolean(DataValidation validation, bool isInputMessage, bool? value)
    {
        BooleanValue? current = isInputMessage ? validation.ShowInputMessage : validation.ShowErrorMessage;
        if (current?.Value == value)
        {
            return false;
        }

        BooleanValue? next = value.HasValue ? new BooleanValue(value.Value) : null;
        if (isInputMessage)
        {
            validation.ShowInputMessage = next;
        }
        else
        {
            validation.ShowErrorMessage = next;
        }

        return true;
    }

    private static (int r1, int c1, int r2, int c2) ParseReferenceArgument(string reference)
    {
        if (TryParseReference(reference, out var bounds))
        {
            return bounds;
        }

        throw new ArgumentException($"Invalid A1 reference '{reference}'.", nameof(reference));
    }

    private static bool ReferenceListOverlaps(string referenceList, (int r1, int c1, int r2, int c2) filter)
    {
        foreach (string reference in referenceList.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
        {
            if (TryParseReference(reference, out var bounds) && RangesOverlapInclusive(filter, bounds))
            {
                return true;
            }
        }

        return false;
    }

    private static bool TryParseReference(string reference, out (int r1, int c1, int r2, int c2) bounds)
    {
        bounds = default;
        if (string.IsNullOrWhiteSpace(reference))
        {
            return false;
        }

        string normalized = reference.Trim();
        int sheetSeparator = normalized.LastIndexOf('!');
        if (sheetSeparator >= 0)
        {
            normalized = normalized.Substring(sheetSeparator + 1);
        }

        string[] parts = normalized.Split(':');
        if (parts.Length > 2)
        {
            return false;
        }

        if (!TryParseCellReference(parts[0], out int r1, out int c1))
        {
            return false;
        }

        int r2 = r1;
        int c2 = c1;
        if (parts.Length == 2 && !TryParseCellReference(parts[1], out r2, out c2))
        {
            return false;
        }

        if (r2 < r1)
        {
            (r1, r2) = (r2, r1);
        }

        if (c2 < c1)
        {
            (c1, c2) = (c2, c1);
        }

        bounds = (r1, c1, r2, c2);
        return true;
    }

    private static bool TryParseCellReference(string reference, out int row, out int column)
    {
        row = 0;
        column = 0;
        if (string.IsNullOrWhiteSpace(reference))
        {
            return false;
        }

        string normalized = reference.Trim().Replace("$", string.Empty);
        int index = 0;
        while (index < normalized.Length && char.IsLetter(normalized[index]))
        {
            column = (column * 26) + (char.ToUpperInvariant(normalized[index]) - 'A' + 1);
            index++;
        }

        if (column == 0 || index == normalized.Length)
        {
            return false;
        }

        for (; index < normalized.Length; index++)
        {
            if (!char.IsDigit(normalized[index]))
            {
                return false;
            }

            row = (row * 10) + (normalized[index] - '0');
        }

        return row > 0;
    }

    private static bool RangesOverlapInclusive((int r1, int c1, int r2, int c2) first, (int r1, int c1, int r2, int c2) second)
        => first.r1 <= second.r2
            && second.r1 <= first.r2
            && first.c1 <= second.c2
            && second.c1 <= first.c2;

    private sealed class DataValidationDisplaySnapshot
    {
        public DataValidationDisplaySnapshot(string range, bool? showInputMessage, bool? showErrorMessage, bool hasInputMessageText, bool hasErrorMessageText)
        {
            Range = range;
            ShowInputMessage = showInputMessage;
            ShowErrorMessage = showErrorMessage;
            HasInputMessageText = hasInputMessageText;
            HasErrorMessageText = hasErrorMessageText;
        }

        public string Range { get; }

        public bool? ShowInputMessage { get; }

        public bool? ShowErrorMessage { get; }

        public bool HasInputMessageText { get; }

        public bool HasErrorMessageText { get; }
    }
}
