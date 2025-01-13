Function Get-PBIAuditLogQuery {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $Id,

        [parameter()]
        [switch]
        $IncludeNonPBIAuditQuery
    )

    $url = "https://graph.microsoft.com/beta/security/auditLog/queries"

    if ($Id) {
        $url = "$url/$($Id)"
    }

    try {
        $response = @(Invoke-MgGraphRequest -Uri $url -OutputType PSObject -ErrorAction Stop)

        if ($response.value) {
            $queryjob = $response.value
        }
        else {
            $queryjob = $response
        }

        if ($IncludeNonPBIAuditQuery) {
            $queryjob
        }
        else {
            $queryjob | Where-Object { $_.recordTypeFilters -eq 'PowerBIAudit' }
        }
    }
    catch {
        $_.Exception.Message | Out-Default
        return $null
    }
}