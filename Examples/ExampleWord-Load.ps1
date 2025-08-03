Import-Module .\PSWriteOffice.psd1 -Force

$File = "C:\Support\GitHub\EFLegacyConfiguration\ScriptsAdHoc\Data\PAM Role ID vs. ADM accounts Compliance.xlsx"

$ImportedData1 = Import-OfficeExcel -FilePath $File
$ImportedData1