
<#PSScriptInfo

.VERSION 1.1

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
    [Parameter(Mandatory)]
    $StartDate,

    [Parameter(Mandatory)]
    $EndDate,

    [Parameter()]
    [int]
    $SplitTimeIntoChunksOf = 1,

    [Parameter()]
    [int]
    $PageSize = 5000,

    [Parameter()]
    [bool]
    $ShowProgress = $true,

    [Parameter()]
    [int]
    $MaxRetryCount = 3
)

if ([datetime]($StartDate) -eq [datetime]$EndDate) {
    "The StartDate and EndDate cannot be the same values." | Write-Verbose
    return $null
}

if ([datetime]($EndDate) -le [datetime]($StartDate)) {
    "The EndDate value cannot be older than the StartDate value." | Write-Verbose
    return $null
}

## Function to Split the search period
Function Split-Time {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)]
        [datetime]
        $BottomDate,

        [Parameter(Mandatory)]
        [datetime]
        $CeilingDate,

        [Parameter()]
        [int]
        $Chunks
    )

    # Calculate duration between start and end dates
    $totalDuration = New-TimeSpan -Start $BottomDate -End $CeilingDate

    # Calculate interval for each chunk
    $interval = $totalDuration.TotalHours / $Chunks

    # Initialize an array to store chunk start and end times
    $chunksArray = @()

    # Start time of the first chunk
    $chunkStartTime = $BottomDate

    # Loop through to generate chunk start and end times
    for ($i = 1; $i -le $Chunks; $i++) {
        # End time of the current chunk
        $chunkEndTime = $chunkStartTime.AddHours($interval)

        # If it's the last chunk, adjust the end time to be the end date
        if ($i -eq $Chunks) {
            $chunkEndTime = $CeilingDate
        }

        # Add chunk start and end times to the array
        $chunksArray += [PSCustomObject]@{
            StartDate = $chunkStartTime
            EndDate   = $chunkEndTime
        }

        # Update start time for next chunk to be one second more than the previous end time
        $chunkStartTime = $chunkEndTime.AddSeconds(1)
    }

    # Output the array of chunk start and end times
    return $chunksArray
}

## Define the session ID and record type to use with the Search-UnifiedAuditLog cmdlet.
$recordType = 'PowerBIAudit'
# $retryCount = 0
# $maxRetryCount = 3

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
    "It looks like you forgot to connect to Remote Exchange PowerShell. You should do that first before asking me to do stuff for you." | Write-Verbose
    Return $null
}
catch {
    "Something is wrong. You can see the error below. You should fix it before asking me to try again." | Write-Verbose
    $_.Exception.Message | Write-Verbose

    Return $null
}
#EndRegion

$searchPeriod = Split-Time -BottomDate $StartDate -CeilingDate $EndDate -Chunks $SplitTimeIntoChunksOf
"The search period will be split into these time chunks." | Write-Verbose
foreach ($period in $searchPeriod) {
    $period | Select-Object StartDate, EndDate | Write-Verbose
}

$searchCounter = 0
foreach ($period in $SearchPeriod) {
    $sessionID = (New-Guid).GUID
    $searchCounter++
    "Search # $($searchCounter) of $($searchPeriod.Count)" | Write-Verbose
    # "Start Date: $($period.StartDate)" | Write-Verbose
    # "End Date: $($period.EndDate)" | Write-Verbose
    "Session Id: $($sessionId)" | Write-Verbose
    $period | Select-Object StartDate, EndDate | Write-Verbose

    do {
        try {
            $currentPageResult = @(Search-UnifiedAuditLog -SessionId $sessionID -SessionCommand ReturnLargeSet -StartDate $period.StartDate -EndDate $period.EndDate -Formatted -RecordType $recordType -ResultSize $PageSize -ErrorAction Stop)
        }
        catch {
            Write-Error "Failed to execute search: $_"
            break
        }

        if ($currentPageResult) {
            $currentPageResult | Add-Member -MemberType NoteProperty -Name SessionId -Value $sessionID
            $currentPageResult
        }
    }
    while ($currentPageResult.Count -gt 0)
}

