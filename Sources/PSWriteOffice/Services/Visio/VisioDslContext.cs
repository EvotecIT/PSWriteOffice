using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Threading;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;

namespace PSWriteOffice.Services.Visio;

internal sealed class VisioDslContext : IDisposable
{
    private static readonly AsyncLocal<VisioDslContext?> CurrentScope = new();
    private readonly Stack<VisioPage> _pages = new();
    private readonly Dictionary<string, VisioShape> _shapesByKey = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, VisioStencilCatalog> _stencilCatalogsByKey = new(StringComparer.OrdinalIgnoreCase);

    private VisioDslContext(VisioDocument document)
    {
        Document = document ?? throw new ArgumentNullException(nameof(document));
    }

    internal VisioDocument Document { get; }

    internal static VisioDslContext? Current => CurrentScope.Value;

    internal static VisioDslContext Enter(VisioDocument document)
    {
        if (CurrentScope.Value != null)
        {
            throw new InvalidOperationException("A Visio DSL scope is already active on this runspace.");
        }

        var scope = new VisioDslContext(document);
        CurrentScope.Value = scope;
        return scope;
    }

    internal static VisioDslContext Require(PSCmdlet caller)
    {
        return CurrentScope.Value ?? throw new PSInvalidOperationException(
            $"'{caller.MyInvocation.InvocationName}' must run inside New-OfficeVisio.");
    }

    internal VisioPage? CurrentPage => _pages.LastOrDefault();

    internal VisioStencilCatalog? DefaultStencilCatalog { get; private set; }

    internal VisioPage RequirePage()
    {
        return CurrentPage ?? throw new PSInvalidOperationException("No Visio page context available. Use VisioPage first.");
    }

    internal IDisposable Push(VisioPage page)
    {
        if (page == null)
        {
            throw new ArgumentNullException(nameof(page));
        }

        _pages.Push(page);
        return new PopToken(this, page);
    }

    internal void RegisterShape(string? key, VisioShape shape)
    {
        if (shape == null)
        {
            throw new ArgumentNullException(nameof(shape));
        }

        if (!string.IsNullOrWhiteSpace(key))
        {
            _shapesByKey[key!] = shape;
        }

        if (!string.IsNullOrWhiteSpace(shape.Name))
        {
            _shapesByKey[shape.Name!] = shape;
        }

        if (!string.IsNullOrWhiteSpace(shape.Id))
        {
            _shapesByKey[shape.Id] = shape;
        }
    }

    internal VisioShape ResolveShape(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new PSArgumentException("Shape reference cannot be empty.", nameof(value));
        }

        if (_shapesByKey.TryGetValue(value, out var shape))
        {
            return shape;
        }

        var page = RequirePage();
        shape = page.AllShapes().FirstOrDefault(candidate =>
            string.Equals(candidate.Id, value, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(candidate.Name, value, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(candidate.NameU, value, StringComparison.OrdinalIgnoreCase));

        if (shape != null)
        {
            RegisterShape(value, shape);
            return shape;
        }

        throw new PSInvalidOperationException($"Visio shape '{value}' was not found in the current DSL scope.");
    }

    internal void RegisterStencilCatalog(string? key, VisioStencilCatalog catalog, bool makeDefault)
    {
        if (catalog == null)
        {
            throw new ArgumentNullException(nameof(catalog));
        }

        if (!string.IsNullOrWhiteSpace(key))
        {
            _stencilCatalogsByKey[key!] = catalog;
        }

        if (!string.IsNullOrWhiteSpace(catalog.Name))
        {
            _stencilCatalogsByKey[catalog.Name] = catalog;
        }

        if (DefaultStencilCatalog == null || makeDefault)
        {
            DefaultStencilCatalog = catalog;
        }
    }

    internal VisioStencilCatalog ResolveStencilCatalog(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return DefaultStencilCatalog ?? OfficeIMO.Visio.Stencils.VisioStencils.All;
        }

        if (_stencilCatalogsByKey.TryGetValue(value!, out var catalog))
        {
            return catalog;
        }

        throw new PSInvalidOperationException($"Visio stencil catalog '{value}' was not found in the current DSL scope.");
    }

    private void Pop(VisioPage page)
    {
        if (_pages.Count > 0 && ReferenceEquals(_pages.Peek(), page))
        {
            _pages.Pop();
        }
    }

    public void Dispose()
    {
        if (CurrentScope.Value == this)
        {
            CurrentScope.Value = null;
        }

        _pages.Clear();
        _shapesByKey.Clear();
        _stencilCatalogsByKey.Clear();
        DefaultStencilCatalog = null;
    }

    private sealed class PopToken : IDisposable
    {
        private VisioDslContext? _context;
        private readonly VisioPage _page;

        internal PopToken(VisioDslContext context, VisioPage page)
        {
            _context = context;
            _page = page;
        }

        public void Dispose()
        {
            _context?.Pop(_page);
            _context = null;
        }
    }
}
