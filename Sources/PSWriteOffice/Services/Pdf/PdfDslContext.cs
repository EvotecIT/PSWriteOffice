using System;
using System.Management.Automation;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Services.Pdf;

internal sealed class PdfDslContext : IDisposable
{
    private static readonly System.Threading.AsyncLocal<PdfDslContext?> Current = new();
    private readonly PdfDslContext? _previous;

    private PdfDslContext(PdfDocument document)
    {
        Document = document;
        _previous = Current.Value;
        Current.Value = this;
    }

    public PdfDocument Document { get; }

    public static PdfDslContext Enter(PdfDocument document) => new(document);

    public static PdfDslContext Require(PSCmdlet cmdlet)
    {
        return Current.Value ?? throw new PSInvalidOperationException(
            $"No active PDF DSL context. Use {cmdlet.MyInvocation.InvocationName} inside New-OfficePdf {{ ... }} or pass -Document.");
    }

    public void Dispose()
    {
        Current.Value = _previous;
    }
}
