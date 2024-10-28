Clear-Host
#Import-Module .\PSWriteOffice.psd1 -Force

$Presentation = New-OfficePowerPoint -FilePath "$PSScriptRoot\Documents\ExamplePowerPoint1.pptx"

$Presentation.Slides.AddEmptySlide([ShapeCrawler.SlideLayoutType]::Title)

Write-Color -Text $Presentation.Slides.Count -Color Green

# Get the shapes collection from the first slide
$shapes = $Presentation.Slides[0].Shapes

# Add a new rectangle shape
$shapes.AddRectangle(50, 60, 100, 70)

$shapes.Item(0).TextBox.Text = "Hello World!"

$Presentation.Slides[1].Shapes | Format-Table

$Presentation.Slides[1].Shapes[0].TextBox.Text = "This is my title"

Save-OfficePowerPoint -Presentation $Presentation -Show