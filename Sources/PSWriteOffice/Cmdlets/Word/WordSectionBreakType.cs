namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Specifies how a new Word section begins.</summary>
public enum WordSectionBreakType
{
    /// <summary>Start the section on the next page.</summary>
    NextPage,

    /// <summary>Start the section in the next column.</summary>
    NextColumn,

    /// <summary>Start the section without forcing a new page.</summary>
    Continuous,

    /// <summary>Start the section on the next even-numbered page.</summary>
    EvenPage,

    /// <summary>Start the section on the next odd-numbered page.</summary>
    OddPage
}
