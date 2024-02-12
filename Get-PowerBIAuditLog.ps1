
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
    $EndDate = (Get-Date)
)

Begin {
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
        Continue
    }
    catch {
        "Something is wrong. You can see the error below. You should fix it before asking me to try again." | Out-Default
        $_.Exception.Message | Out-Default
        Continue
    }
    #EndRegion

    "Start Date: $($StartDate)" | Out-Default
    "End Date: $($EndDate)" | Out-Default

    if ($StartDate -eq $EndDate) {
        "The StartDate and EndDate cannot be the same values." | Out-Default
        Continue
    }

    if ($EndDate -le $StartDate) {
        "The EndDate value cannot be older than the StartDate value." | Out-Default
        Continue
    }

    ## Define the session ID and record type to use with the Search-UnifiedAuditLog cmdlet.
    $sessionID = (New-Guid).GUID
    $recordType = 'PowerBIAudit'
}

process {
    do {
        ## Run the Search-UnifiedAuditLog
        $currentPageResult = Search-UnifiedAuditLog -SessionId $sessionId -SessionCommand ReturnLargeSet -StartDate $startDate -EndDate $endDate -Formatted -RecordType $recordType
        if ($currentPageResult) {
            ## Initialize the maximum results available variable once.
            if (!$maxResultCount) { $maxResultCount = $($currentPageResult[-1].ResultCount) }
            ## Set the current page result count.
            $currentPageResultCount = $($currentPageResult[-1].ResultIndex)
            ## Compute the completion percentage
            $percentComplete = ($currentPageResultCount * 100) / $maxResultCount
            ## Display the progress
            Write-Progress -Activity "Getting Power BI Audit Log [$($StartDate) - $($EndDate)]..." -Status "Result: $($currentPageResultCount) of $($maxResultCount)" -PercentComplete $percentComplete -ErrorAction SilentlyContinue
            ## Display the current page results
            $currentPageResult
        }
    } while (
        ## Continue running while the last ResultIndex in the current page is less than the ResultCount value.
        ## Note: "ResultIndex" is not ZERO-based.
        $currentPageResultCount -lt $maxResultCount
    )
}

end {}




