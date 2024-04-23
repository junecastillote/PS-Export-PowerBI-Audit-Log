# Import-Module ExchangeOnlineManagement
# Connect-ExchangeOnline -AppId "app id" -Organization "orgname.onmicrosoft.com" -CertificateThumbprint "thumbprint"

# $endDate = (Get-Date).AddHours(-18)
# $startDate = $endDate.AddHours(-2)

$endDate = (Get-Date -Hour 0 -Minute 0 -Second 0)
$startDate = (Get-Date $endDate).AddDays(-1)

$outputDirectory = ".\report"
$orgName = (Get-OrganizationConfig).DisplayName
$csvFileName = "$($outputDirectory)\$($orgName -replace ' ','-')_PowerBIAuditLogs_$(Get-Date $startDate -Format "yyyy-dd-MM_H-mm-ss")--$(Get-Date $endDate -Format "yyyy-dd-MM_H-mm-ss").csv"
$zipFileName = $(($csvFileName).Replace('.csv', '.zip'))

$powerbi_logs = .\Get-PowerBIAuditLog.ps1 -StartDate $startDate -EndDate $endDate
$powerbi_logs | Export-Csv -Path $csvFileName -NoTypeInformation -Force
Compress-Archive -Path $csvFileName -DestinationPath $zipFileName -CompressionLevel Optimal -Force
