    $TenantId= "2eea08b8-1972-447b-ad43-d044d042500a"
    $ClientId="62a03bb6-b52a-4b2d-a3a7-82e542026446"
    $ClientSecret="6Xp8Q~PEE3L.1qRNR9EsSdV9.Adxi5mMFy_8TbgC"

$Period = "D180"

$OutputFolder = ".\\output"


Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$ts][$Level] $Message"
}

function Get-AccessToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )

    $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/token"
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        resource      = "https://graph.microsoft.com"
    }

    $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body -ContentType "application/x-www-form-urlencoded"
    return [string]$resp.access_token
}

function Invoke-GraphReportCsv {
    param(
        [string]$Uri,
        [hashtable]$Headers
    )

    try {
        $resp = Invoke-WebRequest -Method Get -Uri $Uri -Headers $Headers -MaximumRedirection 5 -ErrorAction Stop
        return [string]$resp.Content
    } catch {
        $raw = ""
        try { $raw = $_.ErrorDetails.Message } catch {}
        if ([string]::IsNullOrWhiteSpace($raw)) { $raw = $_.Exception.Message }

        if ($raw -match "S2SUnauthorized|Invalid permission|403") {
            throw "Acces refuse Graph Reports. Ajoute la permission Application 'Reports.Read.All' et fais le consentement admin."
        }

        throw
    }
}

try {
    Write-Log "Test tenant-level Graph Reports ($Period)" "INFO"

    if (-not (Test-Path -LiteralPath $OutputFolder)) {
        [void](New-Item -Path $OutputFolder -ItemType Directory -Force)
    }

    $token = Get-AccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    $headers = @{ Authorization = "Bearer $token" }

    # Tenant-level timeseries by service (contains Power BI rows in many tenants)
    $svcUri = "https://graph.microsoft.com/v1.0/reports/getOffice365ServicesUserCounts(period='$Period')"
    $svcCsvText = Invoke-GraphReportCsv -Uri $svcUri -Headers $headers
    $svcRows = @($svcCsvText | ConvertFrom-Csv)

    # Optional user-detail report for cross-check (can be empty depending on tenant/report settings)
    $usrUri = "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='$Period')"
    $usrCsvText = Invoke-GraphReportCsv -Uri $usrUri -Headers $headers
    $usrRows = @($usrCsvText | ConvertFrom-Csv)

    $powerBiSvcRows = @($svcRows | Where-Object {
        $line = ($_ | ConvertTo-Json -Compress)
        $line -match '(?i)Power\s*BI|POWER_BI|Fabric|PBI_'
    })

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $svcAllPath = Join-Path $OutputFolder ("tenant_services_user_counts_{0}_{1}.csv" -f $Period, $stamp)
    $svcPbiPath = Join-Path $OutputFolder ("tenant_services_powerbi_{0}_{1}.csv" -f $Period, $stamp)
    $usrAllPath = Join-Path $OutputFolder ("tenant_active_user_detail_{0}_{1}.csv" -f $Period, $stamp)

    $svcRows | Export-Csv -Path $svcAllPath -NoTypeInformation -Encoding UTF8
    $powerBiSvcRows | Export-Csv -Path $svcPbiPath -NoTypeInformation -Encoding UTF8
    $usrRows | Export-Csv -Path $usrAllPath -NoTypeInformation -Encoding UTF8

    Write-Log ("Services rows total: {0}" -f $svcRows.Length) "SUCCESS"
    Write-Log ("Services rows Power BI/Fabric: {0}" -f $powerBiSvcRows.Length) "SUCCESS"
    Write-Log ("Active user detail rows: {0}" -f $usrRows.Length) "SUCCESS"
    Write-Log ("Export: {0}" -f $svcAllPath) "SUCCESS"
    Write-Log ("Export Power BI: {0}" -f $svcPbiPath) "SUCCESS"
    Write-Log ("Export users: {0}" -f $usrAllPath) "SUCCESS"
}
catch {
    Write-Log ("Erreur: {0}" -f $_.Exception.Message) "ERROR"
    throw
}
