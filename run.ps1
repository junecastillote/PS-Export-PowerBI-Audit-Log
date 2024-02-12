# Import-Module ExchangeOnlineManagement
# Connect-ExchangeOnline -AppId "app id" -Organization "orgname.onmicrosoft.com" -CertificateThumbprint "thumbprint"

$endDate = (Get-Date).AddHours(-18)
$startDate = $endDate.AddHours(-2)
$outputDirectory = ".\report"

.\Get-PowerBIAuditLog.ps1 -StartDate $startDate -EndDate $endDate -OutputDirectory $outputDirectory
