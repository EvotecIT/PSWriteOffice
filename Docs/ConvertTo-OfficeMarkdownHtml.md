---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficeMarkdownHtml
## SYNOPSIS
Converts Markdown content to HTML.

## SYNTAX
### Path (Default)
```powershell
ConvertTo-OfficeMarkdownHtml [-InputPath] <string> [-OutputPath <string>] [-DocumentMode] [-Style <HtmlStyle>] [-CssDelivery <CssDelivery>] [-AssetMode <AssetMode>] [-Title <string>] [-ReaderOptions <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-BaseUri <string>] [-MaxInputCharacters <int>] [-NormalizeInput <MarkdownInputNormalizationPreset>] [-DisallowFileUrls <bool>] [-AllowDataUrls <bool>] [-AllowMailtoUrls <bool>] [-AllowProtocolRelativeUrls <bool>] [-RestrictUrlSchemes <bool>] [-AllowedUrlScheme <string[]>] [-Theme <MarkdownVisualThemeKind>] [-RawHtmlHandling <RawHtmlHandling>] [-IncludeAnchorLinks] [-GitHubTaskListHtml] [-GitHubFootnoteHtml] [-ExternalLinksTargetBlank] [-ExternalLinksRel <string>] [-ExternalLinksReferrerPolicy <string>] [-RestrictHttpLinksToBaseOrigin] [-RestrictHttpImagesToBaseOrigin] [-BlockExternalHttpImages] [-ImagesLoadingLazy] [-ImagesDecodingAsync] [-ImagesReferrerPolicy <string>] [-AllowedHttpLinkHost <string[]>] [-AllowedHttpImageHost <string[]>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Text
```powershell
ConvertTo-OfficeMarkdownHtml -Text <string> [-OutputPath <string>] [-DocumentMode] [-Style <HtmlStyle>] [-CssDelivery <CssDelivery>] [-AssetMode <AssetMode>] [-Title <string>] [-ReaderOptions <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-BaseUri <string>] [-MaxInputCharacters <int>] [-NormalizeInput <MarkdownInputNormalizationPreset>] [-DisallowFileUrls <bool>] [-AllowDataUrls <bool>] [-AllowMailtoUrls <bool>] [-AllowProtocolRelativeUrls <bool>] [-RestrictUrlSchemes <bool>] [-AllowedUrlScheme <string[]>] [-Theme <MarkdownVisualThemeKind>] [-RawHtmlHandling <RawHtmlHandling>] [-IncludeAnchorLinks] [-GitHubTaskListHtml] [-GitHubFootnoteHtml] [-ExternalLinksTargetBlank] [-ExternalLinksRel <string>] [-ExternalLinksReferrerPolicy <string>] [-RestrictHttpLinksToBaseOrigin] [-RestrictHttpImagesToBaseOrigin] [-BlockExternalHttpImages] [-ImagesLoadingLazy] [-ImagesDecodingAsync] [-ImagesReferrerPolicy <string>] [-AllowedHttpLinkHost <string[]>] [-AllowedHttpImageHost <string[]>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
ConvertTo-OfficeMarkdownHtml -Document <MarkdownDoc> [-OutputPath <string>] [-DocumentMode] [-Style <HtmlStyle>] [-CssDelivery <CssDelivery>] [-AssetMode <AssetMode>] [-Title <string>] [-ReaderOptions <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-BaseUri <string>] [-MaxInputCharacters <int>] [-NormalizeInput <MarkdownInputNormalizationPreset>] [-DisallowFileUrls <bool>] [-AllowDataUrls <bool>] [-AllowMailtoUrls <bool>] [-AllowProtocolRelativeUrls <bool>] [-RestrictUrlSchemes <bool>] [-AllowedUrlScheme <string[]>] [-Theme <MarkdownVisualThemeKind>] [-RawHtmlHandling <RawHtmlHandling>] [-IncludeAnchorLinks] [-GitHubTaskListHtml] [-GitHubFootnoteHtml] [-ExternalLinksTargetBlank] [-ExternalLinksRel <string>] [-ExternalLinksReferrerPolicy <string>] [-RestrictHttpLinksToBaseOrigin] [-RestrictHttpImagesToBaseOrigin] [-BlockExternalHttpImages] [-ImagesLoadingLazy] [-ImagesDecodingAsync] [-ImagesReferrerPolicy <string>] [-AllowedHttpLinkHost <string[]>] [-AllowedHttpImageHost <string[]>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Returns HTML text or saves it to a file when -OutputPath is specified.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $html = ConvertTo-OfficeMarkdownHtml -Path .\README.md
```

Returns the rendered HTML.

### EXAMPLE 2
```powershell
PS> ConvertTo-OfficeMarkdownHtml -Path .\Report.md -DocumentMode -Title 'Weekly Report' -Style Clean -OutputPath .\Report.html -PassThru
```

Generates a full HTML file with title and CSS styling.

## PARAMETERS

### -AllowDataUrls
Allow data URLs while parsing Markdown links and images.

```yaml
Type: Nullable`1
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AllowedHttpImageHost
Allowed HTTP(S) image hosts.

```yaml
Type: String[]
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AllowedHttpLinkHost
Allowed HTTP(S) link hosts.

```yaml
Type: String[]
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AllowedUrlScheme
Allowed URL schemes when URL scheme restriction is enabled.

```yaml
Type: String[]
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AllowMailtoUrls
Allow mailto URLs while parsing Markdown links.

```yaml
Type: Nullable`1
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AllowProtocolRelativeUrls
Allow protocol-relative URLs while parsing Markdown links and images.

```yaml
Type: Nullable`1
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AssetMode
Asset loading mode.

```yaml
Type: AssetMode
Parameter Sets: Path, Text, Document
Aliases: None
Possible values: Online, Offline

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -BaseUri
Base URI used to resolve and restrict relative Markdown links and images.

```yaml
Type: String
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -BlockExternalHttpImages
Block all absolute external HTTP(S) images.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CssDelivery
CSS delivery mode.

```yaml
Type: CssDelivery
Parameter Sets: Path, Text, Document
Aliases: None
Possible values: Inline, ExternalFile, LinkHref, None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DisallowFileUrls
Block file URLs while parsing Markdown links and images.

```yaml
Type: Nullable`1
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Markdown document to convert.

```yaml
Type: MarkdownDoc
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -DocumentMode
Render a full HTML document instead of a fragment.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ExternalLinksReferrerPolicy
referrerpolicy value for external HTTP(S) links.

```yaml
Type: String
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ExternalLinksRel
rel attribute value for external HTTP(S) links.

```yaml
Type: String
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ExternalLinksTargetBlank
Open external HTTP(S) links in a new tab.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -GitHubFootnoteHtml
Emit GitHub-compatible footnote HTML.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -GitHubTaskListHtml
Emit GitHub-compatible task-list HTML.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ImagesDecodingAsync
Add decoding="async" to rendered images.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ImagesLoadingLazy
Add loading="lazy" to rendered images.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ImagesReferrerPolicy
referrerpolicy value for rendered images.

```yaml
Type: String
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -IncludeAnchorLinks
Add anchor links to headings.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -InputPath
Path to the Markdown file.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath, Path
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MaxInputCharacters
Maximum Markdown input length accepted by the reader.

```yaml
Type: Nullable`1
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NormalizeInput
Applies a built-in Markdown input normalization preset before parsing.

```yaml
Type: Nullable`1
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OutputPath
Optional output path for the HTML file.

```yaml
Type: String
Parameter Sets: Path, Text, Document
Aliases: OutPath
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit a FileInfo when saving to disk.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Profile
Named reader profile used when ReaderOptions is not supplied.

```yaml
Type: Nullable`1
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -RawHtmlHandling
Controls how raw HTML blocks are emitted.

```yaml
Type: Nullable`1
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ReaderOptions
Optional reader options when parsing Markdown.

```yaml
Type: MarkdownReaderOptions
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -RestrictHttpImagesToBaseOrigin
Restrict absolute HTTP(S) images to the base origin.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -RestrictHttpLinksToBaseOrigin
Restrict absolute HTTP(S) links to the base origin.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -RestrictUrlSchemes
Restrict parsed URL schemes to the allow-list.

```yaml
Type: Nullable`1
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Style
Built-in HTML style preset.

```yaml
Type: HtmlStyle
Parameter Sets: Path, Text, Document
Aliases: None
Possible values: Plain, Clean, GithubLight, GithubDark, GithubAuto, ChatLight, ChatDark, ChatAuto, Word

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Text
Markdown text to convert.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Theme
Shared Markdown visual theme for HTML output.

```yaml
Type: Nullable`1
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Title
Optional title for HTML documents.

```yaml
Type: String
Parameter Sets: Path, Text, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## OUTPUTS

- `System.String
System.IO.FileInfo`

## RELATED LINKS

- None
