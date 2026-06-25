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

    public CsvPowerShellLineWriter(PSCmdlet cmdlet)
    {
        _cmdlet = cmdlet ?? throw new ArgumentNullException(nameof(cmdlet));
    }

    public override Encoding Encoding => Encoding.UTF8;

    public override void Write(char value)
    {
        if (_pendingCarriageReturn)
        {
            if (value == '\n')
            {
                EmitLine();
                _pendingCarriageReturn = false;
                return;
            }

            EmitLine();
            _pendingCarriageReturn = false;
        }

        if (value == '\r')
        {
            _pendingCarriageReturn = true;
            return;
        }

        if (value == '\n')
        {
            EmitLine();
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
                EmitLine();
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
    }
}
