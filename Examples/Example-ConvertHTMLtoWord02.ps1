Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

# This seems to be very limited, as some things just doesn't work
# Small table will work, but large tables won't be processed correctly
$Objects = @(
    [PSCustomObject] @{ Test = 1; Test2 = 'Test'; Test3 = 'Ok' }
    [PSCustomObject] @{ Test = 1; Test2 = 'Test'; Test3 = 'Ok' }
)

$Test = New-HTML {
    New-HTMLText -Text 'This is a test', ' another test' -FontSize 30pt
    New-HTMLTable -DataTable $Objects -simplify
} -Online

ConvertFrom-HTMLToWord -OutputFile $PSScriptRoot\Documents\TestHTML.docx -HTML $Test -Show