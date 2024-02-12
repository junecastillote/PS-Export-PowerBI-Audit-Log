
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

    [Parameter()]
    [string]
    $ExportFileName,

    [Parameter()]
    [bool]
    $ReturnResult = $true
)


## Set progress bar visibility
$ProgressPreference = 'Continue'

## Set progress bar style if PowerShell Core
if ($PSVersionTable.PSEdition -eq 'Core') {
    $PSStyle.Progress.View = 'Classic'
}

#Region - Is Exchange Connected?
try {
    $null = (Get-OrganizationConfig -ErrorAction STOP).DisplayName
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
$script:retryCount = 0
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

## Define the session ID and record type to use with the Search-UnifiedAuditLog cmdlet.
$script:sessionID = (New-Guid).GUID
$script:recordType = 'PowerBIAudit'

## Create or overwrite the export file.
if ($ExportFileName) {
    try {
        $null = New-Item -ItemType File -Force -Path $ExportFileName -ErrorAction Stop
    }
    catch {
        ## If the export file cannot be created, exit the script.
        $_.Exception.Message | Out-Default
        $LASTEXITCODE = 3
        return $null
    }
}


Write-Progress -Activity "Getting Power BI Audit Log [$($StartDate) - $($EndDate)]..." -Status "Progress: Getting the first 100 records (0%)" -PercentComplete 0 -ErrorAction SilentlyContinue
do {
    $currentPageResult = @(ExtractPBILogs)
    if ($currentPageResult) {
        ## Initialize the maximum results available variable once.
        if (!$maxResultCount) {
            $maxResultCount = $($currentPageResult[-1].ResultCount)
            "Total entries: $($maxResultCount)" | Out-Default
        }
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
        if ($ExportFileName) {
            $currentPageResult | Export-Csv -Path $ExportFileName -Append
        }
    }
}
while (
    ## Continue running while the last ResultIndex in the current page is less than the ResultCount value.
    ## Note: "ResultIndex" is not ZERO-based.
        ($currentPageResultCount -lt $maxResultCount) -or ($currentPageResult.Count -gt 0)
)


Write-Progress -Activity "Getting Power BI Audit Log [$($StartDate) - $($EndDate)]..." -Status "Progress: $($currentPageResultCount) of $($maxResultCount) ($([int]$percentComplete)%)" -PercentComplete $percentComplete -ErrorAction SilentlyContinue -Completed

if ($ExportFileName) {
    $csv_file = Get-ChildItem -Path $ExportFileName
    $zip_filename = $(($csv_file.FullName.ToString()).Replace($(((Split-Path $csv_file.FullName -Leaf).Split('.')[-1])), 'zip'))
    $null = Compress-Archive -Path $ExportFileName -DestinationPath $zip_filename -CompressionLevel Optimal -Force
    Start-Sleep -Seconds 2
    $zip_file = Get-ChildItem -Path $zip_filename
    "CSV result: $($csv_file.FullName)" | Out-Default
    "ZIP result: $($zip_file.FullName)" | Out-Default
}
