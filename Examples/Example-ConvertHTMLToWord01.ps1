Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

# This seems to be very limited, as some things just doesn't work
# Small table will work, but large tables won't be processed correctly
$Objects = @(
    [PSCustomObject] @{ Test = 1; Test2 = 'Test'; Test3 = 'Ok' }
    [PSCustomObject] @{ Test = 1; Test2 = 'Test'; Test3 = 'Ok' }
)

New-HTML {
    New-HTMLText -Text 'This is a test', ' another test' -FontSize 30pt
    New-HTMLTable -DataTable $Objects -Simplify
} -Online -FilePath $PSScriptRoot\Documents\Test.html

ConvertFrom-HTMLToWord -OutputFile $PSScriptRoot\Documents\TestHTML.docx -FileHTML $PSScriptRoot\Documents\Test.html -Show