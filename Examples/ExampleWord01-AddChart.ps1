Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Document = New-OfficeWord -FilePath $PSScriptRoot\Documents\BasicDocument.docx

$Document.Settings.FontFamily = 'Times New Roman'

$Document.AddHeadersAndFooters()

New-OfficeWordText -Document $Document -Text 'This is a test, very big test ', 'and this should be bold' -Bold $null, $true -Underline Dash, $null

New-OfficeWordText -Document $Document -Text 'This is a test, very big test', 'ooops' -Color Blue, Gold -Alignment Right

$Paragraph = New-OfficeWordText -Document $Document -Text 'Centered' -Color Blue, Gold -Alignment Center -ReturnObject


$Chart = $Document.AddBarChart()
$Chart.AddCategories(@('Category 1', 'Category 2', 'Category 3'))
$Chart.AddChartBar("Brazil", @(1, 2, 3), 'Red')
$Chart.AddChartBar("Germany", @(2, 3, 4), 'Blue')

$Paragraph = New-OfficeWordText -Document $Document -Text 'Centered' -Color Blue, Gold -Alignment Center -ReturnObject

$Test = [System.Collections.Generic.List[int]]::new()
$Test.Add(10)
$Test.Add(35)
$Test.Add(18)
$Test.Add(23)


$Test2 = [System.Collections.Generic.List[int]]::new()
$Test2.Add(100)
$Test2.Add(1)
$Test2.Add(18)
$Test2.Add(230)

$Chart1 = $Document.AddAreaChart("Area Chart")
$Chart1.AddCategories(@('Category 1', 'Category 2', 'Category 3', 'Category 4'))
$Chart1.AddChartArea("Brazil", $Test, 'AliceBlue')
$Chart1.AddChartArea("Germany", $Test2, 'Red')
$Chart1.AddChartArea("Onet", @(1, 2, 3, 4), 'Green')
$Chart1.AddLegend([DocumentFormat.OpenXml.Drawing.Charts.LegendPositionValues]::Top)

$Paragraph = New-OfficeWordText -Document $Document -Text 'Centered' -Color Blue, Gold -Alignment Center -ReturnObject

$Chart2 = $Document.AddBarChart()
$Chart2.AddCategories(@('Category 1', 'Category 2', 'Category 3'))
$Chart2.AddChartBar("Brazil", @(1, 2, 3), 'Red')
$Chart2.AddChartBar("Germany", @(2, 3, 4), 'Blue')

New-OfficeWordText -Document $Document -Text ' Attached to existing paragraph', ' continue' -Paragraph $Paragraph -Color Blue

Save-OfficeWord -Document $Document -Show