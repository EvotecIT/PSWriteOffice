using System;
using System.Management.Automation;
using System.Text.RegularExpressions;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Sets the background for a Word document.</summary>
/// <example>
///   <summary>Apply a solid background color.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Set-OfficeWordBackground -Color '#f4f7fb'</code>
///   <para>Sets the document background to the provided hex color.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeWordBackground", DefaultParameterSetName = ParameterSetColor)]
[OutputType(typeof(WordDocument))]
public sealed class SetOfficeWordBackgroundCommand : PSCmdlet
{
    private const string ParameterSetColor = "Color";
    private const string ParameterSetImage = "Image";
    private static readonly Regex HexColorPattern = new("^#?[0-9a-fA-F]{6}$", RegexOptions.Compiled);

    /// <summary>Document to update when provided explicitly.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordDocument? Document { get; set; }

    /// <summary>Background color in hex format (#RRGGBB or RRGGBB).</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetColor, Position = 0)]
    public string Color { get; set; } = string.Empty;

    /// <summary>Path to the background image.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetImage, Position = 0)]
    public string ImagePath { get; set; } = string.Empty;

    /// <summary>Optional background image width in pixels.</summary>
    [Parameter(ParameterSetName = ParameterSetImage)]
    public double? Width { get; set; }

    /// <summary>Optional background image height in pixels.</summary>
    [Parameter(ParameterSetName = ParameterSetImage)]
    public double? Height { get; set; }

    /// <summary>Emit the updated document.</summary>
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

        if (ParameterSetName == ParameterSetImage)
        {
            var resolvedPath = ResolvePath(ImagePath);
            document.Background.SetImage(resolvedPath, Width, Height);
        }
        else
        {
            if (!HexColorPattern.IsMatch(Color))
            {
                throw new PSArgumentException("Provide a hex color in #RRGGBB or RRGGBB format.", nameof(Color));
            }

            document.Background.SetColorHex(Color);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }

    private string ResolvePath(string path)
    {
        var providerPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(path);
        return System.IO.Path.IsPathRooted(providerPath)
            ? providerPath
            : System.IO.Path.Combine(SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }
}
