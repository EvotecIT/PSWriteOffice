function New-OfficeWordTableStyle {
    [cmdletBinding()]
    param(

    )

    #[ClosedXML.Excel.XLTableTheme]::GetAllThemes()

    #TableStyle tableStyle = new [TableStyle] { Val = "LightShading-Accent1" };

    $TableStyle = [DocumentFormat.OpenXml.Wordprocessing.TableStyle]::new()
    $TableStyle.Val = "LightShading-Accent1"
    #$TableStyle = [DocumentFormat.OpenXml.Wordprocessing.TableStyle] @{
    #    Val = "LightShading-Accent1"
    #}
    #$TableStyle = [DocumentFormat.OpenXml.Wordprocessing.TableStyle] @{
    #    Val = "TableGrid"
    #}
    , $TableStyle
}