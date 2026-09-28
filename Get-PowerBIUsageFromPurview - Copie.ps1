

    $TenantId= "2eea08b8-1972-447b-ad43-d044d042500a"
    $ClientId="62a03bb6-b52a-4b2d-a3a7-82e542026446"
    $ClientSecret="6Xp8Q~PEE3L.1qRNR9EsSdV9.Adxi5mMFy_8TbgC"

 

   
   $StartTime = (Get-Date).ToUniversalTime().AddDays(-1),

  $EndTime = (Get-Date).ToUniversalTime(),

   
    $OutputFolder = ".\output"


Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS", "DEBUG")]
        [string]$Level = "INFO"
    )

    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$ts][$Level] $Message"
}

function Get-AccessToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$Resource
    )

    $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

    $body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "$Resource/.default"
    }

    $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body -ContentType "application/x-www-form-urlencoded"
    return [string]$resp.access_token
}

function Invoke-ApiRequest {
    param(
        [string]$Method,
        [string]$Uri,
        [hashtable]$Headers
    )

    return Invoke-WebRequest -Method $Method -Uri $Uri -Headers $Headers
}

function Ensure-AuditGeneralSubscription {
    param(
        [string]$BaseUri,
        [hashtable]$Headers
    )

    Write-Log "Verification de l abonnement Audit.General" "INFO"

    $listUri = "$BaseUri/subscriptions/list"
    $resp = Invoke-ApiRequest -Method "GET" -Uri $listUri -Headers $Headers
    $subs = @()

    if (-not [string]::IsNullOrWhiteSpace($resp.Content)) {
        $subs = @($resp.Content | ConvertFrom-Json)
    }

    $alreadySubscribed = $false
    foreach ($s in $subs) {
        if ($s.contentType -eq "Audit.General") {
            $alreadySubscribed = $true
            break
        }
    }

    if ($alreadySubscribed) {
        Write-Log "Abonnement Audit.General deja actif" "SUCCESS"
        return
    }

    $startUri = "$BaseUri/subscriptions/start?contentType=Audit.General"
    [void](Invoke-ApiRequest -Method "POST" -Uri $startUri -Headers $Headers)
    Write-Log "Abonnement Audit.General active" "SUCCESS"
}

function Get-ContentUris {
    param(
        [string]$BaseUri,
        [hashtable]$Headers,
        [datetime]$StartUtc,
        [datetime]$EndUtc
    )

    $startText = $StartUtc.ToString("o")
    $endText = $EndUtc.ToString("o")

    $nextUri = "$BaseUri/subscriptions/content?contentType=Audit.General&startTime=$([uri]::EscapeDataString($startText))&endTime=$([uri]::EscapeDataString($endText))"
    $uris = New-Object System.Collections.Generic.List[object]

    do {
        Write-Log "Lecture des blobs de contenu Purview" "DEBUG"
        $resp = Invoke-ApiRequest -Method "GET" -Uri $nextUri -Headers $Headers

        $chunk = @()
        if (-not [string]::IsNullOrWhiteSpace($resp.Content)) {
            $chunk = @($resp.Content | ConvertFrom-Json)
        }

        foreach ($item in $chunk) {
            $uris.Add($item) | Out-Null
        }

        $headerNext = $resp.Headers["NextPageUri"]
        if ($headerNext -and $headerNext.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($headerNext[0])) {
            $nextUri = [string]$headerNext[0]
        } else {
            $nextUri = $null
        }
    } while ($nextUri)

    return $uris
}

function Get-AuditRecordsFromContentUri {
    param(
        [string]$ContentUri,
        [hashtable]$Headers
    )

    $resp = Invoke-RestMethod -Method Get -Uri $ContentUri -Headers $Headers
    return @($resp)
}

function Ensure-OutputFolder {
    param([string]$Folder)

    if (-not (Test-Path -LiteralPath $Folder)) {
        [void](New-Item -Path $Folder -ItemType Directory -Force)
    }
}

function Test-IsPowerBILicenseEvent {
    param([object]$Record)

    $op = ""
    try { if ($Record.Operation) { $op = [string]$Record.Operation } } catch {}

    $text = ""
    try { $text = ($Record | ConvertTo-Json -Depth 15 -Compress) } catch {}

    $hasLicenseVerb = ($op -match '(?i)license|licence|assign|attribut') -or ($text -match '(?i)license|licence|assignedLicenses|servicePlan')
    if (-not $hasLicenseVerb) { return $false }

    $hasPowerBiMarker = $text -match '(?i)POWER_BI|PowerBI|BI_AZURE_P2|PBI_PREMIUM|FABRIC|M365_LIGHTHOUSE_CUSTOMER_PLAN1'
    return $hasPowerBiMarker
}

function Get-PrincipalFromRecord {
    param([object]$Record)

    $who = ""
    try {
        if ($Record.UserId) { $who = [string]$Record.UserId }
        elseif ($Record.ObjectId) { $who = [string]$Record.ObjectId }
        elseif ($Record.UserKey) { $who = [string]$Record.UserKey }
    } catch {}

    if ([string]::IsNullOrWhiteSpace($who)) { return "<unknown>" }
    return $who
}

function Get-PowerBILicenseTypeFromRecord {
    param([object]$Record)

    $text = ""
    try { $text = ($Record | ConvertTo-Json -Depth 15 -Compress) } catch {}

    if ($text -match '(?i)POWER_BI_PRO|PBI_PRO') { return "Power BI Pro" }
    if ($text -match '(?i)PBI_PREMIUM_PER_USER|PREMIUM_PER_USER|PPU') { return "Power BI Premium Per User" }
    if ($text -match '(?i)BI_AZURE_P2|POWER_BI_PREMIUM|PBI_PREMIUM') { return "Power BI Premium" }
    if ($text -match '(?i)FABRIC|M365_LIGHTHOUSE_CUSTOMER_PLAN1') { return "Fabric/Power BI" }
    if ($text -match '(?i)POWER_BI_STANDARD|PBI_STANDARD|POWER_BI_FREE|PBI_FREE') { return "Power BI Free" }

    $op = ""
    try { if ($Record.Operation) { $op = [string]$Record.Operation } } catch {}
    if ($op -match '(?i)license|licence') { return "Power BI (type inconnu)" }

    return ""
}

try {
    Write-Log "Demarrage extraction usage Power BI depuis Purview" "INFO"

 #   if ($StartTime -ge $EndTime) {
 #       throw "StartTime doit etre strictement inferieur a EndTime"
 #   }

    $startUtc = $StartTime.ToUniversalTime()
    $endUtc = $EndTime.ToUniversalTime()

    Write-Log ("Periode demandee UTC: {0} -> {1}" -f $startUtc.ToString("o"), $endUtc.ToString("o")) "INFO"

    $resource = "https://manage.office.com"
    $baseUri = "https://manage.office.com/api/v1.0/$TenantId/activity/feed"

    $token = Get-AccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -Resource $resource
    $headers = @{ Authorization = "Bearer $token" }

    Ensure-AuditGeneralSubscription -BaseUri $baseUri -Headers $headers

    $contentItems = Get-ContentUris -BaseUri $baseUri -Headers $headers -StartUtc $startUtc -EndUtc $endUtc
    Write-Log ("Blobs Purview trouves: {0}" -f $contentItems.Count) "INFO"

    $allRecords = New-Object System.Collections.Generic.List[object]
    $powerBiRecords = New-Object System.Collections.Generic.List[object]
    $licenseEvents = New-Object System.Collections.Generic.List[object]

    $index = 0
    foreach ($item in $contentItems) {
        $index++
        Write-Log ("Lecture blob {0}/{1}" -f $index, $contentItems.Count) "DEBUG"

        $records = Get-AuditRecordsFromContentUri -ContentUri $item.contentUri -Headers $headers
        foreach ($r in $records) {
            $allRecords.Add($r) | Out-Null
            if ($r.Workload -eq "PowerBI") {
                $powerBiRecords.Add($r) | Out-Null
            }

            if (Test-IsPowerBILicenseEvent -Record $r) {
                $who = Get-PrincipalFromRecord -Record $r
                $licenseType = Get-PowerBILicenseTypeFromRecord -Record $r

                $when = $null
                try { if ($r.CreationTime) { $when = [datetime]$r.CreationTime } } catch {}
                if ($null -eq $when) { $when = Get-Date }

                $licenseEvents.Add([pscustomobject]@{
                    Principal     = $who
                    CreationTime  = $when.ToUniversalTime()
                    Operation     = [string]$r.Operation
                    LicenseType   = $licenseType
                }) | Out-Null
            }
        }
    }

    Write-Log ("Total evenements audits lus: {0}" -f $allRecords.Count) "INFO"
    Write-Log ("Total evenements Power BI: {0}" -f $powerBiRecords.Count) "SUCCESS"

    Ensure-OutputFolder -Folder $OutputFolder
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"

    $rawPath = Join-Path -Path $OutputFolder -ChildPath ("powerbi_usage_raw_{0}.json" -f $stamp)
    $csvPath = Join-Path -Path $OutputFolder -ChildPath ("powerbi_usage_{0}.csv" -f $stamp)
    $summaryPath = Join-Path -Path $OutputFolder -ChildPath ("powerbi_usage_summary_{0}.csv" -f $stamp)
    $licenseStartPath = Join-Path -Path $OutputFolder -ChildPath ("powerbi_license_start_{0}.csv" -f $stamp)

    $powerBiRecords | ConvertTo-Json -Depth 20 | Out-File -FilePath $rawPath -Encoding UTF8

    $powerBiRecords |
        Select-Object
            @{Name = "CreationTime"; Expression = { $_.CreationTime }},
            @{Name = "Operation"; Expression = { $_.Operation }},
            @{Name = "UserId"; Expression = { $_.UserId }},
            @{Name = "Workload"; Expression = { $_.Workload }},
            @{Name = "RecordType"; Expression = { $_.RecordType }},
            @{Name = "ObjectId"; Expression = { $_.ObjectId }},
            @{Name = "ClientIP"; Expression = { $_.ClientIP }},
            @{Name = "ResultStatus"; Expression = { $_.ResultStatus }},
            @{Name = "Activity"; Expression = { $_.Activity }},
            @{Name = "WorkspaceName"; Expression = { $_.WorkSpaceName }},
            @{Name = "DatasetName"; Expression = { $_.DatasetName }},
            @{Name = "ReportName"; Expression = { $_.ReportName }} |
        Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    $powerBiRecords |
        Group-Object -Property Operation |
        Sort-Object -Property Count -Descending |
        Select-Object @{Name = "Operation"; Expression = { $_.Name }}, Count |
        Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8

    $usageByPrincipal = @{}
    foreach ($u in $powerBiRecords) {
        $p = Get-PrincipalFromRecord -Record $u
        $t = $null
        try { if ($u.CreationTime) { $t = [datetime]$u.CreationTime } } catch {}
        if ($null -eq $t) { continue }
        $t = $t.ToUniversalTime()

        if (-not $usageByPrincipal.ContainsKey($p)) {
            $usageByPrincipal[$p] = $t
        } elseif ($t -gt $usageByPrincipal[$p]) {
            $usageByPrincipal[$p] = $t
        }
    }

    $threshold6Months = (Get-Date).ToUniversalTime().AddMonths(-6)

    $licenseStarts = @()
    if ($licenseEvents.Count -gt 0) {
        $licenseStarts = $licenseEvents |
            Group-Object -Property Principal |
            ForEach-Object {
                $first = $_.Group | Sort-Object -Property CreationTime | Select-Object -First 1
                $types = $_.Group |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_.LicenseType) } |
                    Select-Object -ExpandProperty LicenseType -Unique
                $licenseTypeOut = if ($types -and $types.Count -gt 0) { ($types -join " | ") } else { "" }

                $lastUsageOut = ""
                if ($usageByPrincipal.ContainsKey($_.Name)) {
                    $lastUsage = [datetime]$usageByPrincipal[$_.Name]
                    if ($lastUsage -ge $threshold6Months) {
                        $lastUsageOut = $lastUsage.ToString("o")
                    }
                }

                [pscustomobject]@{
                    Principal                  = $_.Name
                    LicenseStartDateTimeUtc    = $first.CreationTime.ToString("o")
                    FirstDetectedOperation     = $first.Operation
                    LicenseType                = $licenseTypeOut
                    LastUsageDateTimeUtc       = $lastUsageOut
                }
            } |
            Sort-Object -Property Principal
    }

    $licenseStarts | Export-Csv -Path $licenseStartPath -NoTypeInformation -Encoding UTF8

    Write-Log ("Fichier brut JSON: {0}" -f $rawPath) "SUCCESS"
    Write-Log ("Fichier detail CSV: {0}" -f $csvPath) "SUCCESS"
    Write-Log ("Fichier resume CSV: {0}" -f $summaryPath) "SUCCESS"
    Write-Log ("Fichier date debut licence Power BI: {0}" -f $licenseStartPath) "SUCCESS"
    Write-Log ("Evenements licence Power BI detectes: {0}" -f $licenseEvents.Count) "INFO"
    Write-Log "Extraction terminee" "SUCCESS"
}
catch {
    Write-Log ("Erreur fatale: {0}" -f $_.Exception.Message) "ERROR"
    throw
}
