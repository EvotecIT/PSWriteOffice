Import-Module .\PSWriteOffice.psd1 -Force

$Path = "C:\Support\GitHub\ASMLBenchmark\Documentation\Windows Baselines 2022.07.29.xlsx"

$ImportedData1 = Import-OfficeExcel -FilePath $Path
$ImportedData1 | Format-Table