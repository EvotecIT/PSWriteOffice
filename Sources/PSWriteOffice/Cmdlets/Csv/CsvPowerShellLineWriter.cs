using System;
using System.IO;
using System.Management.Automation;
using System.Text;

namespace PSWriteOffice.Cmdlets.Csv;

internal sealed class CsvPowerShellLineWriter : TextWriter
{
    private readonly PSCmdlet _cmdlet;
    private readonly StringBuilder _line = new();
    private bool _pendingCarriageReturn;
    private bool _inQuotes;
    private bool _pendingQuoteInQuotedField;

    public CsvPowerShellLineWriter(PSCmdlet cmdlet)
    {
        _cmdlet = cmdlet ?? throw new ArgumentNullException(nameof(cmdlet));
    }

    public override Encoding Encoding => Encoding.UTF8;

    public override void Write(char value)
    {
        if (ResolvePendingQuote(value))
        {
            return;
        }

        if (_pendingCarriageReturn)
        {
            if (value == '\n')
            {
                if (_inQuotes)
                {
                    _line.Append("\r\n");
                }
                else
                {
                    EmitLine();
                }

                _pendingCarriageReturn = false;
                return;
            }

            if (_inQuotes)
            {
                _line.Append('\r');
            }
            else
            {
                EmitLine();
            }

            _pendingCarriageReturn = false;
        }

        if (value == '"')
        {
            _line.Append(value);
            if (_inQuotes)
            {
                _pendingQuoteInQuotedField = true;
            }
            else
            {
                _inQuotes = true;
            }

            return;
        }

        if (value == '\r')
        {
            _pendingCarriageReturn = true;
            return;
        }

        if (value == '\n')
        {
            if (_inQuotes)
            {
                _line.Append(value);
            }
            else
            {
                EmitLine();
            }

            return;
        }

        _line.Append(value);
    }

    public override void Write(string? value)
    {
        if (value is null)
        {
            return;
        }

        foreach (var character in value)
        {
            Write(character);
        }
    }

#if NET6_0_OR_GREATER
    public override void Write(StringBuilder? value)
    {
        if (value is null)
        {
            return;
        }

        for (var i = 0; i < value.Length; i++)
        {
            Write(value[i]);
        }
    }
#endif

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            if (_pendingCarriageReturn)
            {
                if (_inQuotes)
                {
                    _line.Append('\r');
                }
                else
                {
                    EmitLine();
                }

                _pendingCarriageReturn = false;
            }

            if (_line.Length > 0)
            {
                EmitLine();
            }
        }

        base.Dispose(disposing);
    }

    private void EmitLine()
    {
        _cmdlet.WriteObject(_line.ToString());
        _line.Clear();
        _inQuotes = false;
        _pendingQuoteInQuotedField = false;
    }

    private bool ResolvePendingQuote(char value)
    {
        if (!_pendingQuoteInQuotedField)
        {
            return false;
        }

        _pendingQuoteInQuotedField = false;
        if (value == '"')
        {
            _line.Append(value);
            return true;
        }

        _inQuotes = false;
        return false;
    }
}
