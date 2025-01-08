Function Get-PBIAuditLogQuery {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $Id
    )

    $url = "https://graph.microsoft.com/beta/security/auditLog/queries"

    if ($Id) {
        $url = "https://graph.microsoft.com/beta/security/auditLog/queries/$($Id)"
    }

    try {
        $queryjob = @(Invoke-MgGraphRequest -Uri $url -OutputType PSObject -ErrorAction Stop)
        if ($queryjob.value) {
            $queryjob.value | Where-Object { $_.recordTypeFilters -eq 'PowerBIAudit' }
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