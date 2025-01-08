Function Get-PBIAuditLogQueryRecord {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $Id,

        [Parameter()]
        [ValidateRange(30, 300)]
        [int]
        $Wait
    )

    $InformationPreference = 'Continue'
    $WarningPreference = 'Continue'

    if (!$Wait) {
        $queryjob = Get-PBIAuditLogQuery -Id $Id
        if ($queryjob.status -in @("notStarted", "running")) {
            "Query status ($($queryjob.id)): [$($queryjob.status)]" | Write-Warning
            "Wait for the query to complete before re-retrying, or use the -Wait switch." | Write-Error
            return $null
        }
    }

    if ($Wait) {
        do {
            $queryjob = Get-PBIAuditLogQuery -Id $Id
            if ($queryjob.status -notin @("notStarted", "running")) {
                "Query status ($($queryjob.id)): [$($queryjob.status)]" | Write-Information
                break
            }

            "Query status ($($queryjob.id)): [$($queryjob.status)]" | Write-Information
            if ($queryjob.status -in @("notStarted", "running")) {
                "Query status ($($queryjob.id)): [$($queryjob.status)]. Re-checking in $($Wait) seconds." | Write-Information
                Start-Sleep -Seconds $Wait
            }
            else {
                "Query status ($($queryjob.id)): [$($queryjob.status)]" | Write-Information
                break
            }
        }
        while ($true)
    }

    if ($queryjob.status -ne 'succeeded') {
        "The search query ($($queryjob.id)) failed." | Write-Error
        return $null
    }

    $result = [System.Collections.Generic.List[System.Object]]@()
    $page = 0
    $url = "https://graph.microsoft.com/beta/security/auditLog/queries/$($queryjob.id)/records?`$top=1000"

    do {
        $page++
        $response = @(Invoke-MgGraphRequest -Uri $url -Method GET -OutputType PSObject -ErrorAction Stop)

        if ($response.value.Count -lt 1) {
            "No results." | Write-Information
            break
        }

        $result.AddRange($response.value)

        "Result: Page = $page, Count = $($response.value.Count), Total = $($result.Count)" | Write-Information

        $url = $response.'@odata.nextLink'
    }
    while ( $url )
    "Result: Total = $($result.Count)" | Write-Information

    $result
}