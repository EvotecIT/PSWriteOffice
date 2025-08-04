using System;
using System.Management.Automation;

namespace PSWriteOffice.Validation;

public sealed class ValidateScriptAttribute : ValidateEnumeratedArgumentsAttribute
{
    private readonly ScriptBlock _scriptBlock;

    public ValidateScriptAttribute(string script)
    {
        _scriptBlock = ScriptBlock.Create(script ?? throw new ArgumentNullException(nameof(script)));
    }

    protected override void ValidateElement(object element)
    {
        if (element == null)
        {
            throw new ValidationMetadataException("ValidateScriptFailure");
        }

        var result = _scriptBlock.InvokeReturnAsIs(element);
        if (!LanguagePrimitives.IsTrue(result))
        {
            throw new ValidationMetadataException("ValidateScriptFailure");
        }
    }
}
