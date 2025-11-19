using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Threading;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal sealed class ExcelDslContext : IDisposable
{
    private static readonly AsyncLocal<ExcelDslContext?> CurrentScope = new();
    private readonly Stack<object> _scopes = new();

    private ExcelDslContext(ExcelDocument document)
    {
        Document = document ?? throw new ArgumentNullException(nameof(document));
    }

    public ExcelDocument Document { get; }

    public static ExcelDslContext Enter(ExcelDocument document)
    {
        if (CurrentScope.Value != null)
        {
            throw new InvalidOperationException("An Excel DSL scope is already active on this runspace.");
        }

        var context = new ExcelDslContext(document);
        CurrentScope.Value = context;
        return context;
    }

    public static ExcelDslContext Require(PSCmdlet caller)
    {
        var context = CurrentScope.Value;
        if (context == null)
        {
            throw new InvalidOperationException(
                $"'{caller.MyInvocation.InvocationName}' must run inside New-OfficeExcel.");
        }

        return context;
    }

    public IDisposable Push(object scope)
    {
        if (scope == null) throw new ArgumentNullException(nameof(scope));
        _scopes.Push(scope);
        return new PopToken(this, scope);
    }

    private void Pop(object scope)
    {
        if (_scopes.Count == 0)
        {
            return;
        }

        if (ReferenceEquals(_scopes.Peek(), scope))
        {
            _scopes.Pop();
        }
    }

    private sealed class PopToken : IDisposable
    {
        private ExcelDslContext? _context;
        private readonly object _scope;

        public PopToken(ExcelDslContext context, object scope)
        {
            _context = context;
            _scope = scope;
        }

        public void Dispose()
        {
            _context?.Pop(_scope);
            _context = null;
        }
    }

    public ExcelSheet? CurrentSheet => _scopes.OfType<ExcelSheet>().LastOrDefault();

    public ExcelSheet RequireSheet()
    {
        var sheet = CurrentSheet;
        if (sheet == null)
        {
            throw new InvalidOperationException("No worksheet context available. Use Add-OfficeExcelSheet / ExcelSheet first.");
        }

        return sheet;
    }

    public void Dispose()
    {
        if (CurrentScope.Value == this)
        {
            CurrentScope.Value = null;
        }
        _scopes.Clear();
    }
}
