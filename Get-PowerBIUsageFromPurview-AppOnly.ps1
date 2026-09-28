
  
    $TenantId= "2eea08b8-1972-447b-ad43-d044d042500a"
    $ClientId="62a03bb6-b52a-4b2d-a3a7-82e542026446"
    $ClientSecret="6Xp8Q~PEE3L.1qRNR9EsSdV9.Adxi5mMFy_8TbgC"

 
    $StartTime = (Get-Date).ToUniversalTime().AddDays(-90)

    $EndTime = (Get-Date).ToUniversalTime()

    $AutoConnectCompliance = $true
    $ComplianceOrganization = "2eea08b8-1972-447b-ad43-d044d042500a"
    $ComplianceAuthMode = "Credential"
    $ComplianceUserName = "svc_dws@infopro-digital.com"
    $CompliancePassword = "ASawoBQsjMvaKB2"

    
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

function Get-AccessTokenAppOnly {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$Resource
    )

    $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/token"

    $body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        resource      = $Resource
    }

    try {
        $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
        return [string]$resp.access_token
    } catch {
        $raw = ""
        try { $raw = $_.ErrorDetails.Message } catch {}
        if ([string]::IsNullOrWhiteSpace($raw)) {
            try { $raw = $_.Exception.Message } catch {}
        }

        if ($raw -match "AADSTS65001|consent_required") {
            $adminConsentUrl = "https://login.microsoftonline.com/$TenantId/adminconsent?client_id=$ClientId"
            throw "Consentement admin requis pour l application. Ouvre cette URL avec un Global Admin: $adminConsentUrl"
        }

        throw
    }
}

function Invoke-ApiRequest {
    param(
        [string]$Method,
        [string]$Uri,
        [hashtable]$Headers
    )

    try {
        return Invoke-WebRequest -Method $Method -Uri $Uri -Headers $Headers -ErrorAction Stop
    } catch {
        $raw = ""
        try { $raw = $_.ErrorDetails.Message } catch {}
        if ([string]::IsNullOrWhiteSpace($raw)) {
            try { $raw = $_.Exception.Message } catch {}
        }

        if ($raw -match 'AF10001') {
            $msg = @"
Erreur AF10001: le token ne contient pas les permissions attendues pour Office 365 Management API.

Actions a faire dans Entra ID (application):
1) API permissions -> Office 365 Management APIs -> Application permissions
2) Ajouter au minimum: ActivityFeed.Read
3) (Optionnel) Ajouter: ActivityFeed.ReadDlp
4) Cliquer Grant admin consent
5) Attendre 2-5 minutes puis relancer le script
"@
            throw $msg
        }

        throw
    }
}

function ConvertFrom-Base64Url {
    param([string]$Input)

    $s = $Input.Replace('-', '+').Replace('_', '/')
    switch ($s.Length % 4) {
        2 { $s += '==' }
        3 { $s += '=' }
    }
    $bytes = [System.Convert]::FromBase64String($s)
    return [System.Text.Encoding]::UTF8.GetString($bytes)
}

function Get-JwtPayload {
    param([string]$Jwt)

    if ([string]::IsNullOrWhiteSpace($Jwt)) { return $null }
    $parts = $Jwt.Split('.')
    if ($parts.Count -lt 2) { return $null }

    try {
        $json = ConvertFrom-Base64Url -Input $parts[1]
        return ($json | ConvertFrom-Json)
    } catch {
        return $null
    }
}

function Assert-TokenHasManageOfficePermissions {
    param([string]$AccessToken)

    $payload = Get-JwtPayload -Jwt $AccessToken
    if ($null -eq $payload) {
        Write-Log "Impossible de decoder le token pour verifier les roles" "WARN"
        return
    }

    $aud = ""
    try { $aud = [string]$payload.aud } catch {}
    $roles = @()
    try {
        if ($payload.roles) {
            if ($payload.roles -is [System.Array]) { $roles = @($payload.roles) }
            else { $roles = @([string]$payload.roles) }
        }
    } catch {}

    Write-Log ("Token audience: {0}" -f $aud) "DEBUG"
    Write-Log ("Token roles: {0}" -f (($roles -join ', '))) "DEBUG"

    $hasNeededRole = $roles -contains 'ActivityFeed.Read' -or $roles -contains 'ActivityFeed.ReadDlp'
    if (-not $hasNeededRole) {
        $msg = @"
Le token ne contient pas ActivityFeed.Read / ActivityFeed.ReadDlp.

Configure l'application Entra ID:
1) Office 365 Management APIs (Application permissions)
2) Ajouter ActivityFeed.Read
3) Grant admin consent
"@
        throw $msg
    }
}

function Ensure-AuditGeneralSubscription {
    param(
        [string]$BaseUri,
        [hashtable]$Headers
    )

    Write-Log "Verification abonnement Audit.General" "INFO"

    $listUri = "$BaseUri/subscriptions/list"
    $resp = Invoke-ApiRequest -Method "GET" -Uri $listUri -Headers $Headers
    $subs = @()

    if (-not [string]::IsNullOrWhiteSpace($resp.Content)) {
        $parsed = $resp.Content | ConvertFrom-Json

        if ($parsed -is [System.Array]) {
            $subs = @($parsed)
        } else {
            $p = $parsed.PSObject.Properties
            if ($p['subscriptions'] -and $parsed.subscriptions -is [System.Array]) {
                $subs = @($parsed.subscriptions)
            } elseif ($p['content'] -and $parsed.content -is [System.Array]) {
                $subs = @($parsed.content)
            } else {
                $subs = @($parsed)
            }
        }
    }

    $exists = $false
    foreach ($s in $subs) {
        $ct = $null
        if ($null -ne $s -and $s.PSObject -and $s.PSObject.Properties['contentType']) {
            $ct = [string]$s.contentType
        }

        if ($ct -eq "Audit.General") {
            $exists = $true
            break
        }
    }

    if ($exists) {
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

    $uris = New-Object System.Collections.Generic.List[object]

    $nowUtc = (Get-Date).ToUniversalTime()
    $minAllowedStart = $nowUtc.AddDays(-7)
    $safeStart = $StartUtc
    $safeEnd = $EndUtc

    if ($safeEnd -gt $nowUtc) {
        Write-Log "EndTime dans le futur, ajustement a maintenant (UTC)" "WARN"
        $safeEnd = $nowUtc
    }

    if ($safeStart -lt $minAllowedStart) {
        Write-Log ("Limite Office 365 Management API: startTime ne peut pas etre plus ancien que 7 jours. Plage demandee tronquee a {0}." -f $minAllowedStart.ToString("o")) "WARN"
        $safeStart = $minAllowedStart
    }

    if ($safeStart -ge $safeEnd) {
        Write-Log "Aucune fenetre valide a interroger apres ajustement des bornes temporelles" "WARN"
        return $uris
    }

    $windowStart = $safeStart
    while ($windowStart -lt $safeEnd) {
        $windowEnd = $windowStart.AddHours(24)
        if ($windowEnd -gt $safeEnd) {
            $windowEnd = $safeEnd
        }

        $startText = $windowStart.ToString("o")
        $endText = $windowEnd.ToString("o")
        Write-Log ("Interrogation Purview de {0} a {1}" -f $startText, $endText) "DEBUG"

        $nextUri = "$BaseUri/subscriptions/content?contentType=Audit.General&startTime=$([uri]::EscapeDataString($startText))&endTime=$([uri]::EscapeDataString($endText))"

        do {
            $resp = Invoke-ApiRequest -Method "GET" -Uri $nextUri -Headers $Headers

            $chunk = @()
            if (-not [string]::IsNullOrWhiteSpace($resp.Content)) {
                $parsed = $resp.Content | ConvertFrom-Json

                if ($parsed -is [System.Array]) {
                    $chunk = @($parsed)
                } else {
                    $p = $parsed.PSObject.Properties
                    if ($p['content'] -and $parsed.content -is [System.Array]) {
                        $chunk = @($parsed.content)
                    } else {
                        $chunk = @($parsed)
                    }
                }
            }

            foreach ($item in $chunk) {
                $uris.Add($item) | Out-Null
            }

            $headerNext = $resp.Headers["NextPageUri"]
            $nextUriCandidate = $null

            if ($null -ne $headerNext) {
                if ($headerNext -is [System.Array]) {
                    if ($headerNext.Length -gt 0) {
                        $nextUriCandidate = [string]$headerNext[0]
                    }
                } else {
                    $nextUriCandidate = [string]$headerNext
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($nextUriCandidate)) {
                $nextUri = $nextUriCandidate
            } else {
                $nextUri = $null
            }
        } while ($nextUri)

        $windowStart = $windowEnd
    }

    return $uris
}

function Get-AuditRecordsFromContentUri {
    param(
        [string]$ContentUri,
        [hashtable]$Headers
    )

    return @(Invoke-RestMethod -Method Get -Uri $ContentUri -Headers $Headers)
}

function Add-PurviewUsageToCollections {
    param(
        [string]$BaseUri,
        [hashtable]$Headers,
        [datetime]$StartUtc,
        [datetime]$EndUtc,
        [System.Collections.Generic.List[object]]$AllRecords,
        [System.Collections.Generic.List[object]]$PowerBiRecords,
        [System.Collections.Generic.List[object]]$LicenseEvents,
        [string]$LogLabel = "Purview"
    )

    Ensure-AuditGeneralSubscription -BaseUri $BaseUri -Headers $Headers

    $contentItems = Get-ContentUris -BaseUri $BaseUri -Headers $Headers -StartUtc $StartUtc -EndUtc $EndUtc
    Write-Log ("Blobs {0} trouves: {1}" -f $LogLabel, $contentItems.Count) "INFO"

    $index = 0
    foreach ($item in $contentItems) {
        $index++
        Write-Log ("Lecture blob {0} {1}/{2}" -f $LogLabel, $index, $contentItems.Count) "DEBUG"

        $records = Get-AuditRecordsFromContentUri -ContentUri $item.contentUri -Headers $Headers
        foreach ($r in $records) {
            $AllRecords.Add($r) | Out-Null

            if ($r.Workload -eq "PowerBI") {
                $PowerBiRecords.Add($r) | Out-Null
            }

            if (Test-IsPowerBILicenseEvent -Record $r) {
                $who = Get-PrincipalFromRecord -Record $r
                $licenseType = Get-PowerBILicenseTypeFromRecord -Record $r

                $when = $null
                try { if ($r.CreationTime) { $when = [datetime]$r.CreationTime } } catch {}
                if ($null -eq $when) { $when = Get-Date }

                $LicenseEvents.Add([pscustomobject]@{
                    Principal    = $who
                    CreationTime = $when.ToUniversalTime()
                    Operation    = [string]$r.Operation
                    LicenseType  = $licenseType
                }) | Out-Null
            }
        }
    }
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
    if ($text -match '(?i)SPE_E5|SPE_E3|ENTERPRISEPACK|ENTERPRISEPREMIUM|M365_BUSINESS_PREMIUM|M365_BUSINESS_STANDARD|O365_BUSINESS_PREMIUM|STANDARDPACK') { return "Licence Microsoft 365 avec acces Power BI" }

    $op = ""
    try { if ($Record.Operation) { $op = [string]$Record.Operation } } catch {}
    if ($op -match '(?i)license|licence') { return "Power BI (type inconnu)" }

    return ""
}

function Invoke-GraphGetAllPages {
    param(
        [string]$Uri,
        [hashtable]$Headers
    )

    $all = New-Object System.Collections.Generic.List[object]
    $next = $Uri

    while (-not [string]::IsNullOrWhiteSpace($next)) {
        $resp = Invoke-RestMethod -Method Get -Uri $next -Headers $Headers -ErrorAction Stop

        $chunk = @()
        if ($null -ne $resp) {
            $props = $resp.PSObject.Properties
            if ($props['value'] -and $resp.value -is [System.Array]) {
                $chunk = @($resp.value)
            } elseif ($resp -is [System.Array]) {
                $chunk = @($resp)
            } else {
                $chunk = @($resp)
            }
        }

        foreach ($item in $chunk) {
            $all.Add($item) | Out-Null
        }

        $nextCandidate = $null
        if ($null -ne $resp -and $resp.PSObject.Properties['@odata.nextLink']) {
            $nextCandidate = [string]$resp.'@odata.nextLink'
        }

        if ([string]::IsNullOrWhiteSpace($nextCandidate)) {
            $next = $null
        } else {
            $next = $nextCandidate
        }
    }

    return $all
}

function Test-IsPowerBIString {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    return ($Text -match '(?i)POWER_BI|PBI_|BI_AZURE|FABRIC|MICROSOFT_FABRIC')
}

function Test-IsPowerBIAccessSkuPart {
    param([string]$SkuPartNumber)

    if ([string]::IsNullOrWhiteSpace($SkuPartNumber)) { return $false }

    # SKUs commonly granting Power BI access via bundled service plans.
    return ($SkuPartNumber -match '(?i)^SPE_E3$|^SPE_E5$|^ENTERPRISEPACK$|^ENTERPRISEPREMIUM$|^STANDARDPACK$|^O365_BUSINESS_PREMIUM$|^M365_BUSINESS_PREMIUM$|^M365_BUSINESS_STANDARD$|^DEVELOPERPACK$')
}

function Get-PowerBILicensedUsersFromGraph {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )

    Write-Log "Inventaire des licences Power BI via Microsoft Graph" "INFO"

    $graphResource = "https://graph.microsoft.com"
    $graphToken = Get-AccessTokenAppOnly -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -Resource $graphResource
    $headers = @{ Authorization = "Bearer $graphToken" }

    $skuUri = "https://graph.microsoft.com/v1.0/subscribedSkus?`$select=skuId,skuPartNumber,servicePlans"
    $skuRows = Invoke-GraphGetAllPages -Uri $skuUri -Headers $headers

    $skuMap = @{}
    foreach ($s in $skuRows) {
        $skuId = $null
        $skuPart = ""
        $hasPowerBIAccess = $false
        try { if ($s.skuId) { $skuId = [string]$s.skuId } } catch {}
        try { if ($s.skuPartNumber) { $skuPart = [string]$s.skuPartNumber } } catch {}

        try {
            if ($s.servicePlans) {
                foreach ($sp in @($s.servicePlans)) {
                    $planName = ""
                    try { if ($sp.servicePlanName) { $planName = [string]$sp.servicePlanName } } catch {}
                    if (Test-IsPowerBIString -Text $planName) {
                        $hasPowerBIAccess = $true
                        break
                    }
                }
            }
        } catch {}

        if (-not $hasPowerBIAccess) {
            $hasPowerBIAccess = Test-IsPowerBIAccessSkuPart -SkuPartNumber $skuPart
        }

        if (-not [string]::IsNullOrWhiteSpace($skuId)) {
            $skuMap[$skuId] = [pscustomobject]@{
                SkuPart          = $skuPart
                HasPowerBIAccess = $hasPowerBIAccess
            }
        }
    }

    $usersUri = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName,mail,assignedLicenses,assignedPlans&`$top=999"
    $users = Invoke-GraphGetAllPages -Uri $usersUri -Headers $headers

    $out = New-Object System.Collections.Generic.List[object]
    foreach ($u in $users) {
        $principal = ""
        $displayName = ""
        $mail = ""

        try { if ($u.userPrincipalName) { $principal = [string]$u.userPrincipalName } } catch {}
        try { if ($u.displayName) { $displayName = [string]$u.displayName } } catch {}
        try { if ($u.mail) { $mail = [string]$u.mail } } catch {}

        if ([string]::IsNullOrWhiteSpace($principal)) {
            if (-not [string]::IsNullOrWhiteSpace($mail)) { $principal = $mail }
            else {
                try { if ($u.id) { $principal = [string]$u.id } } catch {}
            }
        }
        if ([string]::IsNullOrWhiteSpace($principal)) { continue }

        $assignedSkuIds = New-Object System.Collections.Generic.List[string]
        try {
            if ($u.assignedLicenses) {
                foreach ($al in @($u.assignedLicenses)) {
                    try {
                        if ($al.skuId) {
                            $assignedSkuIds.Add([string]$al.skuId) | Out-Null
                        }
                    } catch {}
                }
            }
        } catch {}

        $licenseLabels = New-Object System.Collections.Generic.List[string]
        foreach ($sid in $assignedSkuIds) {
            if ($skuMap.ContainsKey($sid)) {
                $skuMeta = $skuMap[$sid]
                $part = ""
                $hasPowerBIAccess = $false
                try { if ($skuMeta.SkuPart) { $part = [string]$skuMeta.SkuPart } } catch {}
                try { if ($skuMeta.HasPowerBIAccess) { $hasPowerBIAccess = [bool]$skuMeta.HasPowerBIAccess } } catch {}

                if ($hasPowerBIAccess -or (Test-IsPowerBIString -Text $part)) {
                    $licenseLabels.Add($part) | Out-Null
                }
            }
        }

        $powerBiPlanDates = New-Object System.Collections.Generic.List[datetime]
        try {
            if ($u.assignedPlans) {
                foreach ($ap in @($u.assignedPlans)) {
                    $planName = ""
                    $capStatus = ""
                    try { if ($ap.servicePlanName) { $planName = [string]$ap.servicePlanName } } catch {}
                    try { if ($ap.capabilityStatus) { $capStatus = [string]$ap.capabilityStatus } } catch {}

                    if ((Test-IsPowerBIString -Text $planName) -and ($capStatus -match '(?i)Enabled|Warning|Suspended')) {
                        $licenseLabels.Add($planName) | Out-Null
                        try {
                            if ($ap.assignedDateTime) {
                                $powerBiPlanDates.Add(([datetime]$ap.assignedDateTime).ToUniversalTime()) | Out-Null
                            }
                        } catch {}
                    }
                }
            }
        } catch {}

        $uniqueLabels = @($licenseLabels | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
        if ($uniqueLabels.Length -eq 0) {
            continue
        }

        $graphStart = ""
        if ($powerBiPlanDates.Count -gt 0) {
            $firstGraph = $powerBiPlanDates | Sort-Object | Select-Object -First 1
            $graphStart = $firstGraph.ToString("o")
        }

        $out.Add([pscustomobject]@{
            Principal              = $principal
            DisplayName            = $displayName
            Mail                   = $mail
            GraphLicenseType       = ($uniqueLabels -join " | ")
            GraphLicenseStartUtc   = $graphStart
        }) | Out-Null
    }

    Write-Log ("Utilisateurs avec licence Power BI detectes via Graph: {0}" -f $out.Count) "INFO"
    return $out
}

function Convert-ToIntSafe {
    param([object]$Value)

    if ($null -eq $Value) { return 0 }
    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return 0 }

    $n = 0
    if ([int]::TryParse($s, [ref]$n)) { return $n }
    return 0
}

function Try-OpenComplianceSession {
    param(
        [bool]$AutoConnect,
        [string]$Organization,
        [string]$AuthMode,
        [string]$UserName,
        [string]$Password
    )

    if (-not $AutoConnect) {
        Write-Log "Auto-connexion Compliance desactivee par configuration" "WARN"
        return $false
    }

    $connectCmd = Get-Command -Name "Connect-IPPSSession" -ErrorAction SilentlyContinue
    if ($null -eq $connectCmd) {
        Write-Log "Connect-IPPSSession introuvable dans la session courante" "WARN"
        return $false
    }

    if (($AuthMode -eq "Credential") -and (-not [string]::IsNullOrWhiteSpace($UserName)) -and (-not [string]::IsNullOrWhiteSpace($Password))) {
        try {
            Write-Log "Tentative connexion Compliance en mode identifiant/mot de passe" "INFO"

            $securePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
            $cred = New-Object System.Management.Automation.PSCredential($UserName, $securePassword)

            $credParams = @{ ErrorAction = "Stop"; Credential = $cred }
            if (-not [string]::IsNullOrWhiteSpace($Organization)) {
                $credParams["Organization"] = $Organization
            }
            if ($connectCmd.Parameters.ContainsKey("DisableWAM")) {
                $credParams["DisableWAM"] = $true
            }

            Connect-IPPSSession @credParams | Out-Null
            Write-Log "Session Compliance ouverte en mode identifiant/mot de passe" "SUCCESS"
            return $true
        } catch {
            Write-Log ("Connexion Compliance directe echouee: {0}" -f $_.Exception.Message) "WARN"
        }
    }

    $baseParams = @{ ErrorAction = "Stop" }
    if (-not [string]::IsNullOrWhiteSpace($Organization)) {
        $baseParams["Organization"] = $Organization
    }

    $attempts = New-Object System.Collections.Generic.List[hashtable]
    $attempts.Add($baseParams) | Out-Null

    if ($connectCmd.Parameters.ContainsKey("DisableWAM")) {
        $noWamParams = @{}
        foreach ($k in $baseParams.Keys) { $noWamParams[$k] = $baseParams[$k] }
        $noWamParams["DisableWAM"] = $true
        $attempts.Add($noWamParams) | Out-Null
    }

    if ($connectCmd.Parameters.ContainsKey("UseRPSSession")) {
        $rpsParams = @{}
        foreach ($k in $baseParams.Keys) { $rpsParams[$k] = $baseParams[$k] }
        $rpsParams["UseRPSSession"] = $true
        if ($connectCmd.Parameters.ContainsKey("DisableWAM")) {
            $rpsParams["DisableWAM"] = $true
        }
        $attempts.Add($rpsParams) | Out-Null
    }

    $attemptNumber = 0
    foreach ($p in $attempts) {
        $attemptNumber++

        $mode = "interactive"
        if ($p.ContainsKey("UseRPSSession") -and [bool]$p["UseRPSSession"]) { $mode = "RPS" }
        if ($p.ContainsKey("DisableWAM") -and [bool]$p["DisableWAM"]) { $mode = "$mode + DisableWAM" }

        if (-not [string]::IsNullOrWhiteSpace($Organization)) {
            Write-Log ("Tentative connexion Compliance ({0}/{1}) organisation {2} mode {3}" -f $attemptNumber, $attempts.Count, $Organization, $mode) "INFO"
        } else {
            Write-Log ("Tentative connexion Compliance ({0}/{1}) mode {2}" -f $attemptNumber, $attempts.Count, $mode) "INFO"
        }

        try {
            Connect-IPPSSession @p | Out-Null
            Write-Log "Session Compliance ouverte" "SUCCESS"
            return $true
        } catch {
            $msg = $_.Exception.Message
            Write-Log ("Connexion Compliance echouee (mode {0}): {1}" -f $mode, $msg) "WARN"

            # The WithBroker error is usually fixed by retrying with DisableWAM.
            if ($msg -match "WithBroker" -and -not ($p.ContainsKey("DisableWAM") -and [bool]$p["DisableWAM"])) {
                Write-Log "Erreur Broker detectee, retry automatique en mode DisableWAM" "WARN"
            }
        }
    }

    return $false
}

function Test-SearchUnifiedAuditLogAvailable {
    param(
        [bool]$AutoConnectCompliance,
        [string]$ComplianceOrganization,
        [string]$ComplianceAuthMode,
        [string]$ComplianceUserName,
        [string]$CompliancePassword
    )

    $searchCmd = Get-Command -Name "Search-UnifiedAuditLog" -ErrorAction SilentlyContinue
    if ($null -ne $searchCmd) {
        return $true
    }

    $exoModule = Get-Module -ListAvailable -Name "ExchangeOnlineManagement" | Sort-Object -Property Version -Descending | Select-Object -First 1
    if ($null -eq $exoModule) {
        Write-Log "Module ExchangeOnlineManagement introuvable localement. Search-UnifiedAuditLog necessite une session Purview/Compliance." "WARN"
        return $false
    }

    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop | Out-Null
        Write-Log "Module ExchangeOnlineManagement importe" "DEBUG"
    } catch {
        Write-Log ("Import ExchangeOnlineManagement impossible: {0}" -f $_.Exception.Message) "WARN"
        return $false
    }

    $searchCmd = Get-Command -Name "Search-UnifiedAuditLog" -ErrorAction SilentlyContinue
    if ($null -ne $searchCmd) {
        return $true
    }

    $opened = Try-OpenComplianceSession -AutoConnect $AutoConnectCompliance -Organization $ComplianceOrganization -AuthMode $ComplianceAuthMode -UserName $ComplianceUserName -Password $CompliancePassword
    if ($opened) {
        $searchCmd = Get-Command -Name "Search-UnifiedAuditLog" -ErrorAction SilentlyContinue
        if ($null -ne $searchCmd) {
            return $true
        }
    }

    Write-Log "Search-UnifiedAuditLog reste indisponible. Ouvre une session Purview/Compliance (Connect-IPPSSession ou Connect-ExchangeOnline) avant execution." "WARN"
    return $false
}

function Invoke-PowerBIAuditSearchBatch {
    param(
        [datetime]$StartDate,
        [datetime]$EndDate,
        [string]$SessionId
    )

    try {
        return @(Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -RecordType PowerBI -SessionId $SessionId -SessionCommand ReturnLargeSet -ResultSize 5000 -ErrorAction Stop)
    } catch {
        $msg = ""
        try { $msg = [string]$_.Exception.Message } catch {}
        $details = ""
        try { if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $details = [string]$_.ErrorDetails.Message } } catch {}
        $position = ""
        try { if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) { $position = [string]$_.InvocationInfo.PositionMessage } } catch {}
        if (-not [string]::IsNullOrWhiteSpace($details)) {
            $msg = ("{0} | Details: {1}" -f $msg, $details)
        }
        if (-not [string]::IsNullOrWhiteSpace($position)) {
            $msg = ("{0} | Position: {1}" -f $msg, $position)
        }

        $recordTypeIssue = $false
        if ($msg -match "RecordType|Cannot bind parameter|Cannot convert value|A parameter cannot be found") {
            $recordTypeIssue = $true
        }

        if (-not $recordTypeIssue) {
            throw ("Search-UnifiedAuditLog a echoue (mode RecordType): {0}" -f $msg)
        }

        try {
            return @(Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -FreeText "PowerBI" -SessionId $SessionId -SessionCommand ReturnLargeSet -ResultSize 5000 -ErrorAction Stop)
        } catch {
            $msg2 = ""
            try { $msg2 = [string]$_.Exception.Message } catch {}
            $details2 = ""
            try { if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $details2 = [string]$_.ErrorDetails.Message } } catch {}
            if (-not [string]::IsNullOrWhiteSpace($details2)) {
                $msg2 = ("{0} | Details: {1}" -f $msg2, $details2)
            }
            throw ("Search-UnifiedAuditLog a echoue (fallback FreeText): {0}" -f $msg2)
        }
    }
}

function Get-SearchUnifiedAuditLogWindowIsolated {
    param(
        [datetime]$StartUtc,
        [datetime]$EndUtc,
        [string]$SessionId,
        [string]$Organization,
        [string]$AuthMode,
        [string]$UserName,
        [string]$Password
    )

    $psExe = ""
    try {
        $psCmd = Get-Command -Name "powershell" -ErrorAction SilentlyContinue
        if ($null -ne $psCmd) {
            $psExe = [string]$psCmd.Source
        }
    } catch {}
    if ([string]::IsNullOrWhiteSpace($psExe)) {
        try {
            $pwshCmd = Get-Command -Name "pwsh" -ErrorAction SilentlyContinue
            if ($null -ne $pwshCmd) {
                $psExe = [string]$pwshCmd.Source
            }
        } catch {}
    }
    if ([string]::IsNullOrWhiteSpace($psExe)) {
        throw "Aucun executable PowerShell trouve pour mode isole (pwsh/powershell)."
    }

    $tmpScript = Join-Path -Path $env:TEMP -ChildPath ("pbi_audit_isolated_{0}.ps1" -f ([guid]::NewGuid().ToString("N")))
    $tmpOut = Join-Path -Path $env:TEMP -ChildPath ("pbi_audit_rows_{0}.json" -f ([guid]::NewGuid().ToString("N")))

    $isolateScript = @'
param(
    [string]$StartDateIso,
    [string]$EndDateIso,
    [string]$SessionId,
    [string]$OutFile,
    [string]$Organization,
    [string]$AuthMode,
    [string]$UserName,
    [string]$Password
)

$ErrorActionPreference = "Stop"
try { $PSStyle.OutputRendering = "PlainText" } catch {}

Import-Module ExchangeOnlineManagement -ErrorAction Stop

$connectCmd = Get-Command -Name "Connect-IPPSSession" -ErrorAction SilentlyContinue
if ($null -eq $connectCmd) {
    throw "Connect-IPPSSession introuvable dans le sous-processus."
}

$cp = @{ ErrorAction = "Stop" }
if (-not [string]::IsNullOrWhiteSpace($Organization)) {
    $cp["Organization"] = $Organization
}
if ($connectCmd.Parameters.ContainsKey("DisableWAM")) {
    $cp["DisableWAM"] = $true
}

if (($AuthMode -eq "Credential") -and (-not [string]::IsNullOrWhiteSpace($UserName)) -and (-not [string]::IsNullOrWhiteSpace($Password))) {
    $securePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential($UserName, $securePassword)
    $cp["Credential"] = $cred
}

$attempts = New-Object System.Collections.Generic.List[hashtable]
$attempts.Add($cp) | Out-Null

if ($connectCmd.Parameters.ContainsKey("UseRPSSession")) {
    $cpRps = @{}
    foreach ($k in $cp.Keys) { $cpRps[$k] = $cp[$k] }
    $cpRps["UseRPSSession"] = $true
    $attempts.Add($cpRps) | Out-Null
}

$searchCmd = $null
foreach ($ap in $attempts) {
    $ippsSession = $null
    try {
        $ippsSession = Connect-IPPSSession @ap
    } catch {
        continue
    }

    $searchCmd = Get-Command -Name "Search-UnifiedAuditLog" -ErrorAction SilentlyContinue
    if ($null -eq $searchCmd -and $null -ne $ippsSession) {
        try {
            Import-PSSession -Session $ippsSession -CommandName Search-UnifiedAuditLog -DisableNameChecking -AllowClobber -ErrorAction Stop | Out-Null
            $searchCmd = Get-Command -Name "Search-UnifiedAuditLog" -ErrorAction SilentlyContinue
        } catch {}
    }

    if ($null -ne $searchCmd) {
        break
    }
}

if ($null -eq $searchCmd) {
    $exoConnect = Get-Command -Name "Connect-ExchangeOnline" -ErrorAction SilentlyContinue
    if ($null -ne $exoConnect) {
        $exoParams = @{ ErrorAction = "Stop" }
        if (-not [string]::IsNullOrWhiteSpace($Organization) -and $exoConnect.Parameters.ContainsKey("Organization")) {
            $exoParams["Organization"] = $Organization
        }
        if ($exoConnect.Parameters.ContainsKey("ShowBanner")) {
            $exoParams["ShowBanner"] = $false
        }
        if ($exoConnect.Parameters.ContainsKey("DisableWAM")) {
            $exoParams["DisableWAM"] = $true
        }
        if ($cp.ContainsKey("Credential") -and $exoConnect.Parameters.ContainsKey("Credential")) {
            $exoParams["Credential"] = $cp["Credential"]
        }

        try {
            Connect-ExchangeOnline @exoParams | Out-Null
        } catch {}

        $searchCmd = Get-Command -Name "Search-UnifiedAuditLog" -ErrorAction SilentlyContinue
        if ($null -eq $searchCmd) {
            try {
                $fallbackSession = Get-PSSession | Where-Object { $_.State -eq "Opened" } | Select-Object -First 1
                if ($null -ne $fallbackSession) {
                    Import-PSSession -Session $fallbackSession -CommandName Search-UnifiedAuditLog -DisableNameChecking -AllowClobber -ErrorAction Stop | Out-Null
                    $searchCmd = Get-Command -Name "Search-UnifiedAuditLog" -ErrorAction SilentlyContinue
                }
            } catch {}
        }
    }
}

if ($null -eq $searchCmd) {
    throw "Search-UnifiedAuditLog introuvable dans le sous-processus apres tentatives IPPS/EXO."
}

function Invoke-PowerBIAuditSearchBatch {
    param(
        [datetime]$StartDate,
        [datetime]$EndDate,
        [string]$SessionId
    )

    try {
        return @(Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -RecordType PowerBI -SessionId $SessionId -SessionCommand ReturnLargeSet -ResultSize 5000 -ErrorAction Stop)
    } catch {
        $msg = ""
        try { $msg = [string]$_.Exception.Message } catch {}
        $details = ""
        try { if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $details = [string]$_.ErrorDetails.Message } } catch {}
        $position = ""
        try { if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) { $position = [string]$_.InvocationInfo.PositionMessage } } catch {}
        if (-not [string]::IsNullOrWhiteSpace($details)) {
            $msg = ("{0} | Details: {1}" -f $msg, $details)
        }
        if (-not [string]::IsNullOrWhiteSpace($position)) {
            $msg = ("{0} | Position: {1}" -f $msg, $position)
        }

        $recordTypeIssue = $false
        if ($msg -match "RecordType|Cannot bind parameter|Cannot convert value|A parameter cannot be found") {
            $recordTypeIssue = $true
        }

        if (-not $recordTypeIssue) {
            throw ("Search-UnifiedAuditLog a echoue (mode RecordType): {0}" -f $msg)
        }

        try {
            return @(Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -FreeText "PowerBI" -SessionId $SessionId -SessionCommand ReturnLargeSet -ResultSize 5000 -ErrorAction Stop)
        } catch {
            $msg2 = ""
            try { $msg2 = [string]$_.Exception.Message } catch {}
            $details2 = ""
            try { if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $details2 = [string]$_.ErrorDetails.Message } } catch {}
            if (-not [string]::IsNullOrWhiteSpace($details2)) {
                $msg2 = ("{0} | Details: {1}" -f $msg2, $details2)
            }
            throw ("Search-UnifiedAuditLog a echoue (fallback FreeText): {0}" -f $msg2)
        }
    }
}

$startDate = [datetime]$StartDateIso
$endDate = [datetime]$EndDateIso

$all = New-Object System.Collections.Generic.List[object]
do {
    $batch = @(Invoke-PowerBIAuditSearchBatch -StartDate $startDate -EndDate $endDate -SessionId $SessionId)
    if ($batch.Count -eq 0) {
        break
    }

    foreach ($row in $batch) {
        $all.Add([pscustomobject]@{
            CreationDate = $row.CreationDate
            Operations   = $row.Operations
            UserIds      = $row.UserIds
            Workload     = $row.Workload
            AuditData    = $row.AuditData
        }) | Out-Null
    }
} while ($batch.Count -ge 5000)

$all | ConvertTo-Json -Depth 8 | Out-File -FilePath $OutFile -Encoding UTF8
'@

    $isolateScript | Out-File -FilePath $tmpScript -Encoding UTF8

    try {
        $args = New-Object System.Collections.Generic.List[string]
        $args.Add("-NoProfile") | Out-Null
        $args.Add("-ExecutionPolicy") | Out-Null
        $args.Add("Bypass") | Out-Null
        $args.Add("-File") | Out-Null
        $args.Add($tmpScript) | Out-Null
        $args.Add("-StartDateIso") | Out-Null
        $args.Add($StartUtc.ToString("o")) | Out-Null
        $args.Add("-EndDateIso") | Out-Null
        $args.Add($EndUtc.ToString("o")) | Out-Null
        $args.Add("-SessionId") | Out-Null
        $args.Add($SessionId) | Out-Null
        $args.Add("-OutFile") | Out-Null
        $args.Add($tmpOut) | Out-Null

        if (-not [string]::IsNullOrWhiteSpace($Organization)) {
            $args.Add("-Organization") | Out-Null
            $args.Add([string]$Organization) | Out-Null
        }
        if (-not [string]::IsNullOrWhiteSpace($AuthMode)) {
            $args.Add("-AuthMode") | Out-Null
            $args.Add([string]$AuthMode) | Out-Null
        }
        if (-not [string]::IsNullOrWhiteSpace($UserName)) {
            $args.Add("-UserName") | Out-Null
            $args.Add([string]$UserName) | Out-Null
        }
        if (-not [string]::IsNullOrWhiteSpace($Password)) {
            $args.Add("-Password") | Out-Null
            $args.Add([string]$Password) | Out-Null
        }

        $rawOutput = & $psExe @($args) 2>&1
        if ($LASTEXITCODE -ne 0) {
            $msg = ""
            try { $msg = ($rawOutput | Out-String).Trim() } catch {}
            if ([string]::IsNullOrWhiteSpace($msg)) {
                $msg = "Sous-processus PowerShell en erreur (code $LASTEXITCODE)."
            }
            throw $msg
        }

        if (-not (Test-Path -LiteralPath $tmpOut)) {
            return @()
        }

        $json = Get-Content -LiteralPath $tmpOut -Raw
        if ([string]::IsNullOrWhiteSpace($json)) {
            return @()
        }

        $parsed = $json | ConvertFrom-Json
        if ($parsed -is [System.Array]) {
            return @($parsed)
        }
        return @($parsed)
    } finally {
        try { if (Test-Path -LiteralPath $tmpScript) { Remove-Item -LiteralPath $tmpScript -Force } } catch {}
        try { if (Test-Path -LiteralPath $tmpOut) { Remove-Item -LiteralPath $tmpOut -Force } } catch {}
    }
}

function Get-PowerBIUsageFromAuditLogSearch {
    param(
        [datetime]$StartUtc,
        [datetime]$EndUtc
    )

    Write-Log "Recuperation usage Power BI via Search-UnifiedAuditLog (3 mois)" "INFO"

    $useIsolatedMode = $false
    if (-not (Test-SearchUnifiedAuditLogAvailable -AutoConnectCompliance $AutoConnectCompliance -ComplianceOrganization $ComplianceOrganization -ComplianceAuthMode $ComplianceAuthMode -ComplianceUserName $ComplianceUserName -CompliancePassword $CompliancePassword)) {
        Write-Log "Session locale Compliance indisponible, tentative de collecte en mode isole" "WARN"
        $useIsolatedMode = $true
    }

    $safeStart = $StartUtc.ToUniversalTime()
    $safeEnd = $EndUtc.ToUniversalTime()
    $minStart = (Get-Date).ToUniversalTime().AddDays(-90)
    if ($safeStart -lt $minStart) {
        Write-Log ("Audit Log Search limite a 90 jours dans ce script, StartTime tronque a {0}" -f $minStart.ToString("o")) "WARN"
        $safeStart = $minStart
    }
    if ($safeEnd -le $safeStart) {
        Write-Log "Aucune plage valide pour Audit Log Search" "WARN"
        return @()
    }

    $records = New-Object System.Collections.Generic.List[object]

    $sessionId = "PowerBIUsage_{0}" -f ([guid]::NewGuid().ToString("N"))
    $windowStart = $safeStart
    while ($windowStart -lt $safeEnd) {
        $windowEnd = $windowStart.AddDays(1)
        if ($windowEnd -gt $safeEnd) {
            $windowEnd = $safeEnd
        }

        Write-Log ("Search-UnifiedAuditLog de {0} a {1}" -f $windowStart.ToString("o"), $windowEnd.ToString("o")) "DEBUG"

        if ($useIsolatedMode) {
            $batch = @(Get-SearchUnifiedAuditLogWindowIsolated -StartUtc $windowStart -EndUtc $windowEnd -SessionId $sessionId -Organization $ComplianceOrganization -AuthMode $ComplianceAuthMode -UserName $ComplianceUserName -Password $CompliancePassword)
        } else {
            do {
                $batch = @(Invoke-PowerBIAuditSearchBatch -StartDate $windowStart -EndDate $windowEnd -SessionId $sessionId)
                if ($batch.Count -eq 0) {
                    break
                }

                foreach ($row in $batch) {
                    $auditData = $null
                    try {
                        if (-not [string]::IsNullOrWhiteSpace([string]$row.AuditData)) {
                            $auditData = $row.AuditData | ConvertFrom-Json
                        }
                    } catch {}

                    $creation = ""
                    try {
                        if ($row.CreationDate) {
                            $creation = ([datetime]$row.CreationDate).ToUniversalTime().ToString("o")
                        }
                    } catch {}

                    $operation = ""
                    try { if ($row.Operations) { $operation = [string]$row.Operations } } catch {}
                    if ([string]::IsNullOrWhiteSpace($operation)) {
                        try { if ($auditData.Operation) { $operation = [string]$auditData.Operation } } catch {}
                    }

                    $principal = ""
                    try { if ($row.UserIds) { $principal = [string]$row.UserIds } } catch {}
                    if ([string]::IsNullOrWhiteSpace($principal)) {
                        try { if ($auditData.UserId) { $principal = [string]$auditData.UserId } } catch {}
                    }

                    $workspace = ""
                    try { if ($auditData.WorkSpaceName) { $workspace = [string]$auditData.WorkSpaceName } } catch {}
                    if ([string]::IsNullOrWhiteSpace($workspace)) {
                        try { if ($auditData.WorkspaceName) { $workspace = [string]$auditData.WorkspaceName } } catch {}
                    }

                    $dataset = ""
                    try { if ($auditData.DatasetName) { $dataset = [string]$auditData.DatasetName } } catch {}

                    $report = ""
                    try { if ($auditData.ReportName) { $report = [string]$auditData.ReportName } } catch {}

                    $objectId = ""
                    try { if ($auditData.ObjectId) { $objectId = [string]$auditData.ObjectId } } catch {}

                    $clientIp = ""
                    try { if ($auditData.ClientIP) { $clientIp = [string]$auditData.ClientIP } } catch {}
                    if ([string]::IsNullOrWhiteSpace($clientIp)) {
                        try { if ($auditData.ClientIPAddress) { $clientIp = [string]$auditData.ClientIPAddress } } catch {}
                    }

                    $resultStatus = ""
                    try { if ($auditData.ResultStatus) { $resultStatus = [string]$auditData.ResultStatus } } catch {}

                    $activity = ""
                    try { if ($auditData.Activity) { $activity = [string]$auditData.Activity } } catch {}

                    $records.Add([pscustomobject]@{
                        CreationTime       = $creation
                        Operation          = $operation
                        UserId             = $principal
                        Workload           = "PowerBI"
                        RecordType         = "UnifiedAuditLog"
                        ObjectId           = $objectId
                        ClientIP           = $clientIp
                        ResultStatus       = $resultStatus
                        Activity           = $activity
                        WorkSpaceName      = $workspace
                        DatasetName        = $dataset
                        ReportName         = $report
                        AggregatedActivity = 1
                    }) | Out-Null
                }
            } while ($batch.Count -ge 5000)
        }

        if ($useIsolatedMode) {
            foreach ($row in $batch) {
                $auditData = $null
                try {
                    if (-not [string]::IsNullOrWhiteSpace([string]$row.AuditData)) {
                        $auditData = $row.AuditData | ConvertFrom-Json
                    }
                } catch {}

                $creation = ""
                try {
                    if ($row.CreationDate) {
                        $creation = ([datetime]$row.CreationDate).ToUniversalTime().ToString("o")
                    }
                } catch {}

                $operation = ""
                try { if ($row.Operations) { $operation = [string]$row.Operations } } catch {}
                if ([string]::IsNullOrWhiteSpace($operation)) {
                    try { if ($auditData.Operation) { $operation = [string]$auditData.Operation } } catch {}
                }

                $principal = ""
                try { if ($row.UserIds) { $principal = [string]$row.UserIds } } catch {}
                if ([string]::IsNullOrWhiteSpace($principal)) {
                    try { if ($auditData.UserId) { $principal = [string]$auditData.UserId } } catch {}
                }

                $workspace = ""
                try { if ($auditData.WorkSpaceName) { $workspace = [string]$auditData.WorkSpaceName } } catch {}
                if ([string]::IsNullOrWhiteSpace($workspace)) {
                    try { if ($auditData.WorkspaceName) { $workspace = [string]$auditData.WorkspaceName } } catch {}
                }

                $dataset = ""
                try { if ($auditData.DatasetName) { $dataset = [string]$auditData.DatasetName } } catch {}

                $report = ""
                try { if ($auditData.ReportName) { $report = [string]$auditData.ReportName } } catch {}

                $objectId = ""
                try { if ($auditData.ObjectId) { $objectId = [string]$auditData.ObjectId } } catch {}

                $clientIp = ""
                try { if ($auditData.ClientIP) { $clientIp = [string]$auditData.ClientIP } } catch {}
                if ([string]::IsNullOrWhiteSpace($clientIp)) {
                    try { if ($auditData.ClientIPAddress) { $clientIp = [string]$auditData.ClientIPAddress } } catch {}
                }

                $resultStatus = ""
                try { if ($auditData.ResultStatus) { $resultStatus = [string]$auditData.ResultStatus } } catch {}

                $activity = ""
                try { if ($auditData.Activity) { $activity = [string]$auditData.Activity } } catch {}

                $records.Add([pscustomobject]@{
                    CreationTime       = $creation
                    Operation          = $operation
                    UserId             = $principal
                    Workload           = "PowerBI"
                    RecordType         = "UnifiedAuditLog"
                    ObjectId           = $objectId
                    ClientIP           = $clientIp
                    ResultStatus       = $resultStatus
                    Activity           = $activity
                    WorkSpaceName      = $workspace
                    DatasetName        = $dataset
                    ReportName         = $report
                    AggregatedActivity = 1
                }) | Out-Null
            }
        }

        $windowStart = $windowEnd
    }

    Write-Log ("Enregistrements Power BI recuperes via Audit Log Search: {0}" -f $records.Count) "INFO"
    return $records
}

try {
    Write-Log "Demarrage extraction usage Power BI via Audit Log Search" "INFO"

    if ($StartTime -is [System.Array]) {
        $StartTime = [datetime]$StartTime[0]
        Write-Log "StartTime etait un tableau, premier element utilise" "WARN"
    }
    if ($EndTime -is [System.Array]) {
        $EndTime = [datetime]$EndTime[0]
        Write-Log "EndTime etait un tableau, premier element utilise" "WARN"
    }

    if ($StartTime -ge $EndTime) {
        throw "StartTime doit etre strictement inferieur a EndTime"
    }

    $startUtc = $StartTime.ToUniversalTime()
    $endUtc = $EndTime.ToUniversalTime()
    $powerBiRecords = New-Object System.Collections.Generic.List[object]
    $licenseEvents = New-Object System.Collections.Generic.List[object]

    try {
        $auditUsage = @(Get-PowerBIUsageFromAuditLogSearch -StartUtc $startUtc -EndUtc $endUtc)
        foreach ($u in $auditUsage) {
            $powerBiRecords.Add($u) | Out-Null
        }
    } catch {
        Write-Log ("Collecte via Search-UnifiedAuditLog impossible: {0}" -f $_.Exception.Message) "ERROR"
        throw
    }

    Ensure-OutputFolder -Folder $OutputFolder
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"

    $rawPath = Join-Path -Path $OutputFolder -ChildPath ("powerbi_usage_raw_{0}.json" -f $stamp)
    $csvPath = Join-Path -Path $OutputFolder -ChildPath ("powerbi_usage_{0}.csv" -f $stamp)
    $summaryPath = Join-Path -Path $OutputFolder -ChildPath ("powerbi_usage_summary_{0}.csv" -f $stamp)
    $licenseStartPath = Join-Path -Path $OutputFolder -ChildPath ("powerbi_license_start_{0}.csv" -f $stamp)
    $consolidatedPath = Join-Path -Path $OutputFolder -ChildPath ("powerbi_consolidated_{0}.csv" -f $stamp)

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
    $usageCountByPrincipal = @{}
    foreach ($u in $powerBiRecords) {
        $p = Get-PrincipalFromRecord -Record $u
        $t = $null
        try { if ($u.CreationTime) { $t = [datetime]$u.CreationTime } } catch {}
        if ($null -eq $t) { continue }

        $usageWeight = 1
        try {
            if ($u.PSObject.Properties['AggregatedActivity']) {
                $usageWeight = [int]$u.AggregatedActivity
                if ($usageWeight -lt 1) { $usageWeight = 1 }
            }
        } catch {}

        if (-not $usageCountByPrincipal.ContainsKey($p)) {
            $usageCountByPrincipal[$p] = $usageWeight
        } else {
            $usageCountByPrincipal[$p] = [int]$usageCountByPrincipal[$p] + $usageWeight
        }

        $t = $t.ToUniversalTime()
        if (-not $usageByPrincipal.ContainsKey($p)) {
            $usageByPrincipal[$p] = $t
        } elseif ($t -gt $usageByPrincipal[$p]) {
            $usageByPrincipal[$p] = $t
        }
    }

    $graphLicensedUsers = @()
    try {
        $graphLicensedUsers = @(Get-PowerBILicensedUsersFromGraph -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret)
    } catch {
        Write-Log ("Impossible de recuperer les licences via Graph: {0}" -f $_.Exception.Message) "ERROR"
        Write-Log "Permissions conseillees (Application): User.Read.All, Directory.Read.All, Organization.Read.All (avec consentement admin)." "ERROR"
        throw
    }

    $threshold6Months = (Get-Date).ToUniversalTime().AddMonths(-6)

    $licenseStarts = @()
    if ($licenseEvents.Count -gt 0) {
        $licenseStarts = $licenseEvents |
            Group-Object -Property Principal |
            ForEach-Object {
                $first = $_.Group | Sort-Object -Property CreationTime | Select-Object -First 1
                $types = @($_.Group |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_.LicenseType) } |
                    Select-Object -ExpandProperty LicenseType -Unique)
                $licenseTypeOut = if ($types.Length -gt 0) { ($types -join " | ") } else { "" }

                $lastUsageOut = ""
                if ($usageByPrincipal.ContainsKey($_.Name)) {
                    $lastUsage = [datetime]$usageByPrincipal[$_.Name]
                    if ($lastUsage -ge $threshold6Months) {
                        $lastUsageOut = $lastUsage.ToString("o")
                    }
                }

                [pscustomobject]@{
                    Principal               = $_.Name
                    LicenseStartDateTimeUtc = $first.CreationTime.ToString("o")
                    FirstDetectedOperation  = $first.Operation
                    LicenseType             = $licenseTypeOut
                    LastUsageDateTimeUtc    = $lastUsageOut
                }
            } |
            Sort-Object -Property Principal
    }

    $licenseStarts | Export-Csv -Path $licenseStartPath -NoTypeInformation -Encoding UTF8

    $licenseStartsByPrincipal = @{}
    foreach ($ls in @($licenseStarts)) {
        if ($null -eq $ls) { continue }
        $k = ""
        try { if ($ls.Principal) { $k = [string]$ls.Principal } } catch {}
        if ([string]::IsNullOrWhiteSpace($k)) { continue }
        $licenseStartsByPrincipal[$k] = $ls
    }

    $consolidatedRows = New-Object System.Collections.Generic.List[object]
    foreach ($g in @($graphLicensedUsers)) {
        $principal = [string]$g.Principal
        $lastUsageOut = ""
        $activityCount = 0
        if ($usageByPrincipal.ContainsKey($principal)) {
            $lastUsage = [datetime]$usageByPrincipal[$principal]
            if ($lastUsage -ge $threshold6Months) {
                $lastUsageOut = $lastUsage.ToString("o")
            }
        }
        if ($usageCountByPrincipal.ContainsKey($principal)) {
            $activityCount = [int]$usageCountByPrincipal[$principal]
        }

        $auditStart = ""
        $auditOp = ""
        $auditType = ""
        if ($licenseStartsByPrincipal.ContainsKey($principal)) {
            $audit = $licenseStartsByPrincipal[$principal]
            try { if ($audit.LicenseStartDateTimeUtc) { $auditStart = [string]$audit.LicenseStartDateTimeUtc } } catch {}
            try { if ($audit.FirstDetectedOperation) { $auditOp = [string]$audit.FirstDetectedOperation } } catch {}
            try { if ($audit.LicenseType) { $auditType = [string]$audit.LicenseType } } catch {}
        }

        $licenseStartOut = if (-not [string]::IsNullOrWhiteSpace($auditStart)) { $auditStart } else { [string]$g.GraphLicenseStartUtc }
        $licenseTypeOut = if (-not [string]::IsNullOrWhiteSpace($auditType)) { $auditType } else { [string]$g.GraphLicenseType }

        $consolidatedRows.Add([pscustomobject]@{
            Principal               = $principal
            DisplayName             = [string]$g.DisplayName
            Mail                    = [string]$g.Mail
            HasPowerBILicense       = $true
            LicenseType             = $licenseTypeOut
            LicenseStartDateTimeUtc = $licenseStartOut
            FirstDetectedOperation  = $auditOp
            LastUsageDateTimeUtc    = $lastUsageOut
            PowerBIActivityCount    = $activityCount
            HadPowerBIActivity      = ($activityCount -gt 0)
        }) | Out-Null
    }

    foreach ($principal in @($usageByPrincipal.Keys)) {
        if ([string]::IsNullOrWhiteSpace([string]$principal)) { continue }
        $exists = $false
        foreach ($r in $consolidatedRows) {
            if ([string]$r.Principal -eq [string]$principal) {
                $exists = $true
                break
            }
        }
        if ($exists) { continue }

        $lastUsage = [datetime]$usageByPrincipal[$principal]
        $lastUsageOut = if ($lastUsage -ge $threshold6Months) { $lastUsage.ToString("o") } else { "" }
        $activityCount = 0
        if ($usageCountByPrincipal.ContainsKey($principal)) {
            $activityCount = [int]$usageCountByPrincipal[$principal]
        }

        $auditStart = ""
        $auditOp = ""
        $auditType = ""
        if ($licenseStartsByPrincipal.ContainsKey($principal)) {
            $audit = $licenseStartsByPrincipal[$principal]
            try { if ($audit.LicenseStartDateTimeUtc) { $auditStart = [string]$audit.LicenseStartDateTimeUtc } } catch {}
            try { if ($audit.FirstDetectedOperation) { $auditOp = [string]$audit.FirstDetectedOperation } } catch {}
            try { if ($audit.LicenseType) { $auditType = [string]$audit.LicenseType } } catch {}
        }

        $consolidatedRows.Add([pscustomobject]@{
            Principal               = [string]$principal
            DisplayName             = ""
            Mail                    = ""
            HasPowerBILicense       = $false
            LicenseType             = $auditType
            LicenseStartDateTimeUtc = $auditStart
            FirstDetectedOperation  = $auditOp
            LastUsageDateTimeUtc    = $lastUsageOut
            PowerBIActivityCount    = $activityCount
            HadPowerBIActivity      = ($activityCount -gt 0)
        }) | Out-Null
    }

    $consolidatedRows |
        Sort-Object -Property Principal |
        Export-Csv -Path $consolidatedPath -NoTypeInformation -Encoding UTF8

    Write-Log ("Fichier brut JSON: {0}" -f $rawPath) "SUCCESS"
    Write-Log ("Fichier detail CSV: {0}" -f $csvPath) "SUCCESS"
    Write-Log ("Fichier resume CSV: {0}" -f $summaryPath) "SUCCESS"
    Write-Log ("Fichier date debut/type licence/derniere utilisation: {0}" -f $licenseStartPath) "SUCCESS"
    Write-Log ("Fichier consolide licences + usage: {0}" -f $consolidatedPath) "SUCCESS"
    Write-Log "Extraction terminee" "SUCCESS"
}
catch {
    Write-Log ("Erreur fatale: {0}" -f $_.Exception.Message) "ERROR"
    throw
}
