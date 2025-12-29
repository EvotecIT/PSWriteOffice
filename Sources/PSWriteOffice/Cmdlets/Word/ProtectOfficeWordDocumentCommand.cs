using System;
using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Protects a Word document with a password.</summary>
/// <para>Sets the protection password and the protection type (default: ReadOnly).</para>
/// <example>
///   <summary>Protect a document as read-only.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Protect-OfficeWordDocument -Password 'secret'</code>
///   <para>Applies read-only protection to the current document.</para>
/// </example>
[Cmdlet(VerbsSecurity.Protect, "OfficeWordDocument")]
public sealed class ProtectOfficeWordDocumentCommand : PSCmdlet
{
    /// <summary>Document to protect when provided explicitly.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordDocument? Document { get; set; }

    /// <summary>Password to apply.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Password { get; set; } = string.Empty;

    /// <summary>Protection type (defaults to ReadOnly).</summary>
    [Parameter]
    public DocumentProtectionValues ProtectionType { get; set; } = DocumentProtectionValues.ReadOnly;

    /// <summary>Emit the protected document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = Document ?? WordDslContext.Require(this).Document;
        if (document == null)
        {
            throw new InvalidOperationException("Word document was not provided.");
        }

        document.Settings.ProtectionPassword = Password;
        document.Settings.ProtectionType = ProtectionType;

        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
