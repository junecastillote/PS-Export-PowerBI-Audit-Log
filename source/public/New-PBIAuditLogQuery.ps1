Function New-PBIAuditLogQuery {
    [CmdletBinding()]
    param (
        # Parameter help description
        [Parameter(Mandatory)]
        [datetime]
        $StartDate,

        [Parameter(Mandatory)]
        [datetime]
        $EndDate,

        [Parameter()]
        [string]
        $SearchName = ("PowerBIAudit Search {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm'))
    )

    $Uri = "https://graph.microsoft.com/beta/security/auditLog/queries"
    # $SearchName = ("Audit Search {0}" -f (Get-Date -Format 'dd-MMM-yyyy HH:mm'))
    $SearchParameters = @{
        "@odata.type"                 = "#microsoft.graph.security.auditLogQuery"
        "displayName"                 = $SearchName
        "filterStartDateTime"         = $startDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
        "filterEndDateTime"           = $endDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
        "recordTypeFilters"           = @("PowerBIAudit")
        "keywordFilter"               = ""
        "serviceFilters"              = @()
        "operationFilters"            = @()
        "userPrincipalNameFilters"    = @()
        "ipAddressFilters"            = @()
        "objectIdFilters"             = @()
        "administrativeUnitIdFilters" = @()
        "status"                      = ""
    }

    try {
        $SearchQuery = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $SearchParameters -ContentType 'application/json' -ErrorAction Stop
        return $SearchQuery
    }
    catch {
        $_.Exception.Message | Out-Default
        return $null
    }
}