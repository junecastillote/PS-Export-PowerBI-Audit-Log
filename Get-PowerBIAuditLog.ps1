
<#PSScriptInfo

.VERSION 0.1

.GUID 387f77e0-7781-42a2-8d0b-005580ae6cc4

.AUTHOR June Castillote

.COMPANYNAME

.COPYRIGHT june.castillote@gmail.com

.TAGS

.LICENSEURI

.PROJECTURI https://github.com/junecastillote/PS-Export-PowerBI-Audit-Log

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#>

#Requires -Module ExchangeOnlineManagement

<#

.DESCRIPTION
 PowerShell script wrapper to export large PowerBI audit logs.

#>
[CmdletBinding()]
param (
    [Parameter()]
    [DateTime]
    $StartDate = (Get-Date).AddHours(-24),

    [Parameter()]
    [DateTime]
    $EndDate = (Get-Date),

    [Parameter(Mandatory)]
    [string]
    $OutputDirectory,

    [Parameter()]
    [switch]
    $ReturnResult
)

## Define the session ID and record type to use with the Search-UnifiedAuditLog cmdlet.
$script:sessionID = (New-Guid).GUID
$script:recordType = 'PowerBIAudit'

$script:retryCount = 0
$script:maxRetryCount = 0

## Set progress bar visibility
$ProgressPreference = 'Continue'

## Set progress bar style if PowerShell Core
if ($PSVersionTable.PSEdition -eq 'Core') {
    $PSStyle.Progress.View = 'Classic'
}

#Region - Is Exchange Connected?
try {
    $orgName = (Get-OrganizationConfig -ErrorAction STOP).DisplayName
}
catch [System.Management.Automation.CommandNotFoundException] {
    "It looks like you forgot to connect to Remote Exchange PowerShell. You should do that first before asking me to do stuff for you." | Out-Default
    $LASTEXITCODE = 1
    Return $null
}
catch {
    "Something is wrong. You can see the error below. You should fix it before asking me to try again." | Out-Default
    $_.Exception.Message | Out-Default
    $LASTEXITCODE = 1
    Return $null
}
#EndRegion

#Region ExtractPBILogs
Function ExtractPBILogs {
    Search-UnifiedAuditLog -SessionId $script:sessionID -SessionCommand ReturnLargeSet -StartDate $startDate -EndDate $endDate -Formatted -RecordType $script:recordType
}

#EndRegion

"Start Date: $($StartDate)" | Out-Default
"End Date: $($EndDate)" | Out-Default

if ($StartDate -eq $EndDate) {
    "The StartDate and EndDate cannot be the same values." | Out-Default
    $LASTEXITCODE = 2
    return $null
}

if ($EndDate -le $StartDate) {
    "The EndDate value cannot be older than the StartDate value." | Out-Default
    $LASTEXITCODE = 2
    return $null
}

## Create or overwrite the export file.
if ($OutputDirectory) {
    ## Test if the output directory exists
    if (!(Test-Path $OutputDirectory)) {
        "The output directory [$($OutputDirectory)] does not exist." | Out-Default
        return $null
    }

    $csv_filename = "$($OutputDirectory)\$($orgName -replace ' ','-')_PowerBIAuditLogs_$(Get-Date $startDate -Format "yyyy-dd-MM_H-mm-ss")--$(Get-Date $endDate -Format "yyyy-dd-MM_H-mm-ss").csv"

    try {
        $null = New-Item -ItemType File -Force -Path $csv_filename -ErrorAction Stop
    }
    catch {
        ## If the export file cannot be created, exit the script.
        $_.Exception.Message | Out-Default
        $LASTEXITCODE = 3
        return $null
    }
}

Function IsResultProblematic {
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        $inputObject
    )
    if ($inputObject[-1].ResultIndex -eq -1 -and $inputObject[-1].ResultCount -eq 0) {
        return $true
    }
    else {
        return $false
    }
}



Write-Progress -Activity "Getting Power BI Audit Log [$($StartDate) - $($EndDate)]..." -Status "Progress: Getting the first 100 records (0%)" -PercentComplete 0 -ErrorAction SilentlyContinue
do {
    $currentPageResult = @(ExtractPBILogs)

    if ($currentPageResult.Count -lt 1) {
        "No results found" | Out-Default
        return $null
    }

    if ($script:retryCount -gt $script:maxRetryCount) {
        "The result's total count and indexes are problematic after two retries. This may be a temporary error. Try again after a few minutes." | Out-Default
        return $null
    }

    if (($script:isProblematic = IsResultProblematic -inputObject $currentPageResult) -and ($script:retryCount -le 2)) {
        $script:retryCount++
        $script:sessionID = (New-Guid).Guid
        "Retry # $($script:retryCount++)" | Out-Default
    }
}
while ($script:isProblematic)

## Initialize the maximum results available variable once.
$maxResultCount = $($currentPageResult[-1].ResultCount)
"Total entries: $($maxResultCount)" | Out-Default

## Set the current page result count.
$currentPageResultCount = $($currentPageResult[-1].ResultIndex)
## Compute the completion percentage
$percentComplete = ($currentPageResultCount * 100) / $maxResultCount
## Display the progress
Write-Progress -Activity "Getting Power BI Audit Log [$($StartDate) - $($EndDate)]..." -Status "Progress: $($currentPageResultCount) of $($maxResultCount) ($([int]$percentComplete)%)" -PercentComplete $percentComplete -ErrorAction SilentlyContinue

## Retrieve the rest of the audit log entries
do {
    $currentPageResult = @(ExtractPBILogs)
    if ($currentPageResult) {
        ## Set the current page result count.
        $currentPageResultCount = $($currentPageResult[-1].ResultIndex)
        ## Compute the completion percentage
        $percentComplete = ($currentPageResultCount * 100) / $maxResultCount
        ## Display the progress
        Write-Progress -Activity "Getting Power BI Audit Log [$($StartDate) - $($EndDate)]..." -Status "Progress: $($currentPageResultCount) of $($maxResultCount) ($([int]$percentComplete)%)" -PercentComplete $percentComplete -ErrorAction SilentlyContinue
        ## Display the current page results
        if ($ReturnResult) {
            $currentPageResult
        }
        ## Export result to file
        if ($OutputDirectory) {
            $currentPageResult | Export-Csv -Path $csv_filename -Append
        }
    }
}
while (
    ## Continue running while the last ResultIndex in the current page is less than the ResultCount value.
    ## Note: "ResultIndex" is not ZERO-based.
        ($currentPageResultCount -lt $maxResultCount) -or ($currentPageResult.Count -gt 0)
)

Write-Progress -Activity "Getting Power BI Audit Log [$($StartDate) - $($EndDate)]..." -Status "Progress: $($currentPageResultCount) of $($maxResultCount) ($([int]$percentComplete)%)" -PercentComplete $percentComplete -ErrorAction SilentlyContinue -Completed

if ($OutputDirectory) {
    $csv_file = Get-ChildItem -Path $csv_filename
    $zip_filename = $(($csv_file.FullName.ToString()).Replace('.csv', '.zip'))

    "Compressing results file..." | Out-Default
    $null = Compress-Archive -Path $csv_filename -DestinationPath $zip_filename -CompressionLevel Optimal -Force
    Start-Sleep -Seconds 2

    $zip_file = Get-ChildItem -Path $zip_filename
    "CSV result: $($csv_file.FullName)" | Out-Default
    "ZIP result: $($zip_file.FullName)" | Out-Default
}
