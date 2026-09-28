# =========================
# CONFIGURATION
# =========================

# App Registration TENANT SOURCE
$app_source = @{
tenant_Id = "85a5a352-25d6-4894-a97d-221cd1712dd2"
        client_Id = "28d46b2c-b527-4abb-9fb8-013656682b19"
        client_Secret = "l-f8Q~-NfTX__dizvu8u1dAWSzwlqBcM~unzpcQa"
}

# App Registration TENANT CIBLE
#$app_cible = @{
#    tenant_id     = "2eea08b8-1972-447b-ad43-d044d042500a"
#    client_id     = " 821109a4-6e0e-48e8-b477-8f9b70aa32a4"
#    client_secret = "Yhy8Q~mBxTrYqx~ve7enqTQatjwhEZ5.34-q-a4j"
#}


# App Registration TENANT CIBLE
$app_cible = @{
    tenant_id     = "2eea08b8-1972-447b-ad43-d044d042500a"
    client_id     = " 821109a4-6e0e-48e8-b477-8f9b70aa32a4"
    client_secret = "7DF9F9E37A879A4C043058E849F037173EFF3D54"  # laisser vide si utilisation du certificat
    # Authentification par certificat (recommande pour SharePoint REST)
    # Renseigner certificate_thumbprint (certificat dans le store Windows Cert:\CurrentUser\My)
    # OU certificate_path (chemin vers fichier .pfx)
    certificate_thumbprint = ""  # ex : "A1B2C3D4E5F6..."
    certificate_path       = "C:\certs\MSGraphExchangeOnlineAuth20260717_2.pfx"  # ex : "C:\certs\app_cible.pfx"
    certificate_password   = "MotDePasseFort123!"  # mot de passe du PFX si besoin
}


# =========================
# CHEMIN DU FICHIER CSV
# Format attendu du CSV (séparateur ;) :
#   onedrive_url_source;onedrive_url_cible;upn_source;upn_cible
# Exemple :
#   https://ipdlab-my.sharepoint.com/personal/user01_ipd-lab_com/_layouts/15/onedrive.aspx;https://infoprodigital365-my.sharepoint.com/personal/testmig01_infopro-digital_com/_layouts/15/onedrive.aspx;user01@ipd-lab.com;testmig01@infopro-digital.fr
# =========================
$CsvPath = "$PSScriptRoot\migrations.csv"

# CSV de mapping des utilisateurs source → destination pour restaurer l'auteur des versions
# Format (séparateur ;) : email_source;email_cible
# Exemple : user01@ipd-lab.com;user01@infopro-digital.fr
$UserMappingCsvPath = "$PSScriptRoot\user_mapping.csv"

# Options
$CopyMode = "full"   # full|incremental
$ForceOverwrite = $true
$ChunkSizeMB    = 10
$CopyVersionHistory = $true
$RepairVersionHistoryOnly = $false
$RepairVersionPathRegex = ""   # optionnel, ex: "(?i)^source/repos/"
$AutoRepairVersionHistoryInIncremental = $true
$MaxVersionsPerFile = 0
$ReplayWhenVersionCountMatches = $false
$PruneReplayDuplicateVersions = $false
$PreserveFileTimestamps = $true
$ExportVersionMetadata  = $true
$CopyPermissions        = $true
$CopyItemMetadata       = $true
$EnableEditorIdFallback = $true
$RestoreItemAuthorEditor = $true
$LogLevel               = "INFO"  # DEBUG|INFO|SUCCESS|WARN|ERROR
$AuthorRestoreSkipPathRegex = '(?i)(^|/)(\.git|node_modules|obj|bin|packagetmp)(/|$)'
$AuthorRestoreMaxPathLength = 220
$EnableResumeCheckpoint = $true
$CheckpointEveryItems   = 50
$ResumeStatePath        = "$PSScriptRoot\OneDrive_ResumeState.json"
$EnablePeriodicCsvFlush = $true
$CsvFlushEveryItems     = 200
$UseTempFileTransfer    = $true
$TempTransferDir        = "$env:TEMP\OneDriveCopyTemp"

# Déconnexion préventive du SDK MgGraph pour éviter les conflits de token
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

# =========================
# VARIABLES GLOBALES
# =========================
$script:RunId      = (Get-Date -Format "yyyy-MM-dd_HHmmss")
$script:LogPath    = "$PSScriptRoot\OneDrive_Copy_$($script:RunId).log"
$script:ErrCsvPath = "$PSScriptRoot\OneDrive_Errors_$($script:RunId).csv"
$script:VersionCsvPath = "$PSScriptRoot\OneDrive_Versions_$($script:RunId).csv"
$script:ErrRows    = New-Object System.Collections.Generic.List[object]
$script:VersionRows = New-Object System.Collections.Generic.List[object]
$script:TokenCache = @{}
$script:TokenProfiles = @{}
$script:TokenAliases  = @{}
$script:DstSpToken  = $null
$script:DstSiteInfo = $null
$script:UserMapping = @{}  # source_email -> destination_email
$script:CurrentSourceUpn = ""
$script:CurrentTargetUpn = ""
$script:FolderExistsCache = @{}
$script:EnsureUserIdCache = @{}
$script:ResumeState = @{}
$script:LastPeriodicErrExportCount = 0
$script:LastPeriodicVersionExportCount = 0
$script:CopyModeResolved = "full"
$script:RepairVersionHistoryOnlyResolved = $false

Add-Type -AssemblyName System.Net.Http
$script:HttpClient         = [System.Net.Http.HttpClient]::new()
$script:HttpClient.Timeout = [TimeSpan]::FromMinutes(30)

# =========================
# LOGGING
# =========================
function Write-Log {
    param(
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG')]
        [string]$Level = 'INFO',
        [string]$Message
    )

    $rank = @{ DEBUG = 10; INFO = 20; SUCCESS = 25; WARN = 30; ERROR = 40 }
    $min = if ($rank.ContainsKey($LogLevel)) { $rank[$LogLevel] } else { 20 }
    if ($rank[$Level] -lt $min) { return }

    $ts       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[{0}] [{1}] {2}" -f $ts, $Level, $Message
    switch ($Level) {
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green  }
        "ERROR"   { Write-Host $logEntry -ForegroundColor Red    }
        "WARN"    { Write-Host $logEntry -ForegroundColor Yellow }
        "INFO"    { Write-Host $logEntry -ForegroundColor Cyan   }
        "DEBUG"   { Write-Host $logEntry -ForegroundColor Gray   }
        default   { Write-Host $logEntry -ForegroundColor White  }
    }
    Add-Content -Path $script:LogPath -Value $logEntry -Encoding UTF8
}

function Get-FolderCacheKey {
    param([string]$DriveId, [string]$FolderRelPath)
    $p = if ($FolderRelPath) { ($FolderRelPath -replace '\\', '/').Trim('/') } else { "" }
    return ($DriveId + "|" + $p).ToLower()
}

function Append-Err {
    param([string]$FilePath, [string]$Msg, [string]$Phase = '')
    $script:ErrRows.Add([pscustomobject]@{
        horodatage = Get-Date; fichier = $FilePath; phase = $Phase; erreur = $Msg
    }) | Out-Null
}

function Append-VersionMeta {
    param(
        [string]$FilePath,
        [string]$VersionId,
        [string]$ModifiedUtc,
        [string]$ModifiedBy
    )

    $script:VersionRows.Add([pscustomobject]@{
        horodatage_collecte = Get-Date
        fichier             = $FilePath
        version_id          = $VersionId
        version_modifiee_utc= $ModifiedUtc
        version_modifiee_par= $ModifiedBy
    }) | Out-Null
}

function Test-ShouldSkipAuthorRestore {
    param(
        [string]$ItemRelPath,
        [switch]$ForVersion
    )

    $p = if ($ItemRelPath) { ($ItemRelPath -replace '\\', '/').Trim('/') } else { "" }
    if ([string]::IsNullOrWhiteSpace($p)) { return $false }

    if ($AuthorRestoreMaxPathLength -gt 0 -and $p.Length -gt $AuthorRestoreMaxPathLength) {
        return $true
    }
    if ($AuthorRestoreSkipPathRegex -and $p -match $AuthorRestoreSkipPathRegex) {
        return $true
    }
    # Visual Studio temporary recovery files generate heavy/noisy author restore failures.
    if ($p -match '(?i)~vs[0-9a-f]{3,}\.') {
        return $true
    }
    if ($ForVersion -and $p -match '(?i)(^|/)\.(git|vs)(/|$)') {
        return $true
    }
    return $false
}

function Write-AuthorRestoreIssue {
    param(
        [string]$ItemRelPath,
        [string]$Text,
        [string]$RawMessage = ""
    )

    $probe = if ($RawMessage) { $RawMessage } else { $Text }
    $isKnown = $false
    if ($probe -match 'maxUrlLength|InvalidClientQueryException|404 - File or directory not found|SP REST 401 : \{\}|SP REST 404') {
        $isKnown = $true
    }
    if (Test-ShouldSkipAuthorRestore -ItemRelPath $ItemRelPath) {
        $isKnown = $true
    }

    if ($isKnown) {
        Write-Log "DEBUG" $Text
    } else {
        Write-Log "WARN" $Text
    }
}

function Flush-PeriodicCsvSnapshots {
    param([int]$ProcessedItems)

    if (-not $EnablePeriodicCsvFlush) { return }
    if ($CsvFlushEveryItems -le 0) { return }
    if ($ProcessedItems -le 0) { return }
    if (($ProcessedItems % $CsvFlushEveryItems) -ne 0) { return }

    try {
        if ($script:ErrRows.Count -gt 0 -and $script:ErrRows.Count -ne $script:LastPeriodicErrExportCount) {
            $script:ErrRows | Export-Csv -LiteralPath $script:ErrCsvPath -Encoding UTF8 -NoTypeInformation
            $script:LastPeriodicErrExportCount = $script:ErrRows.Count
        }
        if ($ExportVersionMetadata -and $script:VersionRows.Count -gt 0 -and $script:VersionRows.Count -ne $script:LastPeriodicVersionExportCount) {
            $script:VersionRows | Export-Csv -LiteralPath $script:VersionCsvPath -Encoding UTF8 -NoTypeInformation
            $script:LastPeriodicVersionExportCount = $script:VersionRows.Count
        }
    } catch {
        Write-Log "DEBUG" "Snapshot CSV periodique echoue : $($_.Exception.Message)"
    }
}

function Get-ResumeKey {
    param([string]$UpnSource, [string]$UpnCible)
    return ((Normalize-Email -Value $UpnSource) + "|" + (Normalize-Email -Value $UpnCible)).ToLower()
}

function Load-ResumeState {
    if (-not $EnableResumeCheckpoint) { return }
    $script:ResumeState = @{}
    if (-not (Test-Path -LiteralPath $ResumeStatePath)) { return }
    try {
        $raw = Get-Content -LiteralPath $ResumeStatePath -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return }
        $obj = $raw | ConvertFrom-Json
        foreach ($prop in $obj.PSObject.Properties) {
            $k = [string]$prop.Name
            $v = $prop.Value
            $script:ResumeState[$k] = @{
                FolderIndex = [int]$v.FolderIndex
                FileIndex   = [int]$v.FileIndex
                Updated     = [string]$v.Updated
            }
        }
    } catch {
        Write-Log "WARN" "Checkpoint illisible, ignore : $($_.Exception.Message)"
        $script:ResumeState = @{}
    }
}

function Save-ResumeState {
    if (-not $EnableResumeCheckpoint) { return }
    try {
        $json = ($script:ResumeState | ConvertTo-Json -Depth 8)
        Set-Content -LiteralPath $ResumeStatePath -Value $json -Encoding UTF8
    } catch {
        Write-Log "WARN" "Impossible d'ecrire le checkpoint : $($_.Exception.Message)"
    }
}

function Get-ResumeProgress {
    param([string]$ResumeKey)
    if (-not $EnableResumeCheckpoint) { return $null }
    if ($ResumeKey -and $script:ResumeState.ContainsKey($ResumeKey)) {
        return $script:ResumeState[$ResumeKey]
    }
    return $null
}

function Set-ResumeProgress {
    param(
        [string]$ResumeKey,
        [int]$FolderIndex,
        [int]$FileIndex
    )
    if (-not $EnableResumeCheckpoint -or -not $ResumeKey) { return }
    $script:ResumeState[$ResumeKey] = @{
        FolderIndex = $FolderIndex
        FileIndex   = $FileIndex
        Updated     = (Get-Date).ToString("o")
    }
}

function Clear-ResumeProgress {
    param([string]$ResumeKey)
    if (-not $EnableResumeCheckpoint -or -not $ResumeKey) { return }
    if ($script:ResumeState.ContainsKey($ResumeKey)) {
        $null = $script:ResumeState.Remove($ResumeKey)
    }
}

function Resolve-ActiveToken {
    param([string]$Token)
    if (-not $Token) { return $Token }
    $seen = @{}
    $cur  = $Token
    while ($script:TokenAliases.ContainsKey($cur) -and -not $seen.ContainsKey($cur)) {
        $seen[$cur] = $true
        $cur = [string]$script:TokenAliases[$cur]
    }
    return $cur
}

function Register-TokenProfile {
    param([string]$Token, [hashtable]$Profile)
    if (-not [string]::IsNullOrWhiteSpace($Token) -and $Profile) {
        $script:TokenProfiles[$Token] = $Profile
    }
}

function Refresh-TokenFromProfile {
    param([string]$Token)

    $active = Resolve-ActiveToken -Token $Token
    if (-not $active) { return $null }
    if (-not $script:TokenProfiles.ContainsKey($active)) { return $null }

    $p = $script:TokenProfiles[$active]
    $newToken = $null

    switch ([string]$p.kind) {
        'Graph' {
            $newToken = Get-GraphToken -TenantId $p.tenant_id -ClientId $p.client_id -ClientSecret $p.client_secret -ForceRefresh
        }
        'App' {
            $newToken = Get-AppToken -TenantId $p.tenant_id -ClientId $p.client_id -Scope $p.scope `
                -ClientSecret $p.client_secret -CertThumbprint $p.cert_thumbprint -CertPath $p.cert_path `
                -CertPassword $p.cert_password -Label $p.label -ForceRefresh
        }
        'SharePoint' {
            $newToken = Get-SharePointToken -TenantId $p.tenant_id -ClientId $p.client_id -ClientSecret $p.client_secret -SharePointHost $p.sharepoint_host -ForceRefresh
        }
    }

    if ($newToken) {
        $script:TokenAliases[$Token]  = $newToken
        $script:TokenAliases[$active] = $newToken
        Register-TokenProfile -Token $newToken -Profile $p
    }
    return $newToken
}

# =========================
# TOKEN GRAPH
# =========================

# Build a signed JWT for client_assertion (certificate-based auth)
function New-ClientAssertion {
    param([string]$TenantId, [string]$ClientId, [System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert)
    $x5t     = [System.Convert]::ToBase64String($Cert.GetCertHash()) -replace '\+','-' -replace '/','_' -replace '=',''
    $now     = [int][System.DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $header  = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes('{"alg":"RS256","typ":"JWT","x5t":"' + $x5t + '"}')) -replace '\+','-' -replace '/','_' -replace '=',''
    $payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes('{"aud":"https://login.microsoftonline.com/' + $TenantId + '/oauth2/v2.0/token","iss":"' + $ClientId + '","sub":"' + $ClientId + '","jti":"' + [System.Guid]::NewGuid() + '","nbf":' + $now + ',"exp":' + ($now + 600) + '}')) -replace '\+','-' -replace '/','_' -replace '=',''
    $toSign  = "$header.$payload"
    $data    = [System.Text.Encoding]::UTF8.GetBytes($toSign)

    # GetRSAPrivateKey is an extension method — must be called statically in PS 5.1
    $rsa = $null
    try { $rsa = [System.Security.Cryptography.RSACertificateExtensions]::GetRSAPrivateKey($Cert) } catch {}
    if (-not $rsa) { $rsa = $Cert.PrivateKey }
    if (-not $rsa) { throw "Impossible d'acceder a la cle privee du certificat" }

    if ($rsa -is [System.Security.Cryptography.RSACryptoServiceProvider]) {
        # Re-import into provider type 24 (PROV_RSA_AES) which supports SHA256; type 1 (PROV_RSA_FULL) does not
        $cspParams = [System.Security.Cryptography.CspParameters]::new(24)
        $cspParams.KeyContainerName = [System.Guid]::NewGuid().ToString()
        $rsaAes   = [System.Security.Cryptography.RSACryptoServiceProvider]::new($cspParams)
        $rsaAes.ImportParameters($rsa.ExportParameters($true))
        $sha256   = [System.Security.Cryptography.SHA256Managed]::new()
        $hash     = $sha256.ComputeHash($data)
        $sha256.Dispose()
        $oid      = [System.Security.Cryptography.CryptoConfig]::MapNameToOID("SHA256")
        $sigBytes = $rsaAes.SignHash($hash, $oid)
        $rsaAes.Dispose()
    } else {
        $sigBytes = $rsa.SignData($data, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    }
    $sig = [System.Convert]::ToBase64String($sigBytes) -replace '\+','-' -replace '/','_' -replace '=',''
    return "$toSign.$sig"
}

# Load certificate from store (thumbprint) or PFX file
function Get-AppCertificate {
    param([string]$Thumbprint, [string]$CertPath, [string]$CertPassword)
    if ($Thumbprint) {
        $c = Get-ChildItem "Cert:\CurrentUser\My\$Thumbprint" -ErrorAction SilentlyContinue
        if (-not $c) { $c = Get-ChildItem "Cert:\LocalMachine\My\$Thumbprint" -ErrorAction SilentlyContinue }
        if (-not $c) { throw "Certificat introuvable dans le store (thumbprint=$Thumbprint)" }
        return $c
    }
    if ($CertPath) {
        # Exportable required so the private key is accessible for signing in PS 5.1
        $flags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
        $pfxPassword = if ($CertPassword) { $CertPassword } else { [string]::Empty }
        $cert  = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertPath, $pfxPassword, $flags)
        if (-not $cert.HasPrivateKey) { throw "Le fichier PFX ne contient pas de cle privee : $CertPath" }
        return $cert
    }
    throw "Aucun certificat configure (certificate_thumbprint ou certificate_path)"
}

# Unified token function: uses certificate if configured, otherwise client_secret
function Get-AppToken {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$Scope,
        [string]$ClientSecret       = "",
        [string]$CertThumbprint     = "",
        [string]$CertPath           = "",
        [string]$CertPassword       = "",
        [string]$Label              = "",
        [switch]$ForceRefresh
    )
    $cacheKey = "$TenantId|$ClientId|$Scope"
    $now      = Get-Date
    if (-not $ForceRefresh -and $script:TokenCache.ContainsKey($cacheKey)) {
        $entry = $script:TokenCache[$cacheKey]
        if ($entry.expires -gt $now.Add([TimeSpan]::FromMinutes(5))) { return $entry.token }
    }
    $tokenUri = "https://login.microsoftonline.com/" + $TenantId + "/oauth2/v2.0/token"
    $useCert  = $CertThumbprint -or $CertPath
    if ($useCert) {
        $cert     = Get-AppCertificate -Thumbprint $CertThumbprint -CertPath $CertPath -CertPassword $CertPassword
        $jwt      = New-ClientAssertion -TenantId $TenantId -ClientId $ClientId -Cert $cert
        $bodyStr  = "client_id="                + [Uri]::EscapeDataString($ClientId) +
                    "&client_assertion_type=" + [Uri]::EscapeDataString("urn:ietf:params:oauth:client-assertion-type:jwt-bearer") +
                    "&client_assertion="      + [Uri]::EscapeDataString($jwt) +
                    "&grant_type=client_credentials" +
                    "&scope="                 + [Uri]::EscapeDataString($Scope)
    } else {
        $bodyStr  = "client_id="     + [Uri]::EscapeDataString($ClientId) +
                    "&client_secret=" + [Uri]::EscapeDataString($ClientSecret) +
                    "&grant_type=client_credentials" +
                    "&scope="         + [Uri]::EscapeDataString($Scope)
    }
    try {
        $req         = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Post, $tokenUri)
        $req.Content = [System.Net.Http.StringContent]::new($bodyStr, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
        $resp        = $script:HttpClient.SendAsync($req).Result
        $bodyText    = $resp.Content.ReadAsStringAsync().Result
        $json        = $bodyText | ConvertFrom-Json
        $token       = $json.access_token
        if (-not $token) { throw "Token vide. Reponse : $bodyText" }
        $script:TokenCache[$cacheKey] = @{ token = $token; expires = (Get-Date).AddSeconds([int]$json.expires_in) }
        Register-TokenProfile -Token $token -Profile @{
            kind            = 'App'
            tenant_id       = $TenantId
            client_id       = $ClientId
            scope           = $Scope
            client_secret   = $ClientSecret
            cert_thumbprint = $CertThumbprint
            cert_path       = $CertPath
            cert_password   = $CertPassword
            label           = $Label
        }
        $method = if ($useCert) { "certificat" } else { "secret" }
        if ($Label) { Write-Log "SUCCESS" "Token obtenu ($method) pour $Label" }
        return $token
    } catch {
        throw "Echec token pour tenant $TenantId : $($_.Exception.Message)"
    }
}

function Get-GraphToken {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret,
        [switch]$ForceRefresh
    )
    $cacheKey = "$TenantId|$ClientId|GRAPH"
    $now      = Get-Date
    $skew     = [TimeSpan]::FromMinutes(5)

    if (-not $ForceRefresh -and $script:TokenCache.ContainsKey($cacheKey)) {
        $entry = $script:TokenCache[$cacheKey]
        if ($entry.expires -gt $now.Add($skew)) { return $entry.token }
    }

    $tokenUri = "https://login.microsoftonline.com/" + $TenantId + "/oauth2/v2.0/token"
    $bodyStr  = "client_id="     + [Uri]::EscapeDataString($ClientId) +
                "&client_secret=" + [Uri]::EscapeDataString($ClientSecret) +
                "&grant_type=client_credentials" +
                "&scope="         + [Uri]::EscapeDataString("https://graph.microsoft.com/.default")

    try {
        $req         = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Post, $tokenUri)
        $req.Content = [System.Net.Http.StringContent]::new($bodyStr, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
        $resp        = $script:HttpClient.SendAsync($req).Result
        $bodyText    = $resp.Content.ReadAsStringAsync().Result
        $json        = $bodyText | ConvertFrom-Json
        $token       = $json.access_token

        if (-not $token) { throw "Token vide. Reponse : $bodyText" }

        $script:TokenCache[$cacheKey] = @{
            token   = $token
            expires = (Get-Date).AddSeconds([int]$json.expires_in)
        }
        Register-TokenProfile -Token $token -Profile @{
            kind          = 'Graph'
            tenant_id     = $TenantId
            client_id     = $ClientId
            client_secret = $ClientSecret
        }
        Write-Log "SUCCESS" "Token Graph obtenu pour tenant $TenantId"
        return $token
    } catch {
        Write-Log "ERROR" "Echec token Graph pour $TenantId : $($_.Exception.Message)"
        throw
    }
}

# =========================
# TOKEN SHAREPOINT REST
# =========================
function Get-SharePointToken {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret,
        [Parameter(Mandatory)][string]$SharePointHost,
        [switch]$ForceRefresh
    )
    $scope    = $SharePointHost.TrimEnd('/') + "/.default"
    $cacheKey = "$TenantId|$ClientId|SP|$SharePointHost"
    $now      = Get-Date
    if (-not $ForceRefresh -and $script:TokenCache.ContainsKey($cacheKey)) {
        $entry = $script:TokenCache[$cacheKey]
        if ($entry.expires -gt $now.Add([TimeSpan]::FromMinutes(5))) { return $entry.token }
    }
    $tokenUri = "https://login.microsoftonline.com/" + $TenantId + "/oauth2/v2.0/token"
    $bodyStr  = "client_id="     + [Uri]::EscapeDataString($ClientId) +
                "&client_secret=" + [Uri]::EscapeDataString($ClientSecret) +
                "&grant_type=client_credentials" +
                "&scope="         + [Uri]::EscapeDataString($scope)
    try {
        $req         = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Post, $tokenUri)
        $req.Content = [System.Net.Http.StringContent]::new($bodyStr, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
        $resp        = $script:HttpClient.SendAsync($req).Result
        $bodyText    = $resp.Content.ReadAsStringAsync().Result
        $json        = $bodyText | ConvertFrom-Json
        $token       = $json.access_token
        if (-not $token) { throw "Token vide. Reponse : $bodyText" }
        $script:TokenCache[$cacheKey] = @{ token = $token; expires = (Get-Date).AddSeconds([int]$json.expires_in) }
        Register-TokenProfile -Token $token -Profile @{
            kind           = 'SharePoint'
            tenant_id      = $TenantId
            client_id      = $ClientId
            client_secret  = $ClientSecret
            sharepoint_host= $SharePointHost
        }
        return $token
    } catch {
        throw "Echec token SharePoint pour $SharePointHost : $($_.Exception.Message)"
    }
}

# =========================
# APPEL GRAPH AVEC RETRY
# =========================
function Invoke-GraphCall {
    param(
        [string]$Method      = "GET",
        [string]$Uri,
        [string]$Token,
        [byte[]]$ByteBody    = $null,
        [string]$StringBody  = $null,
        [string]$ContentType = "application/json",
        [int]$MaxRetries     = 5,
        [switch]$ReturnBytes
    )
    $authToken = Resolve-ActiveToken -Token $Token
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        $req = [System.Net.Http.HttpRequestMessage]::new(
            [System.Net.Http.HttpMethod]::new($Method), $Uri)
        $req.Headers.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $authToken)

        if ($ByteBody) {
            $content = [System.Net.Http.ByteArrayContent]::new($ByteBody)
            $content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::new($ContentType)
            $req.Content = $content
        } elseif ($StringBody) {
            $req.Content = [System.Net.Http.StringContent]::new($StringBody, [System.Text.Encoding]::UTF8, $ContentType)
        }

        try {
            $resp     = $script:HttpClient.SendAsync($req).Result
            $bodyText = $resp.Content.ReadAsStringAsync().Result

            if ($ReturnBytes) {
                if ($resp.IsSuccessStatusCode) { return $resp.Content.ReadAsByteArrayAsync().Result }
            } else {
                if ($resp.IsSuccessStatusCode) {
                    if ($bodyText) { return $bodyText | ConvertFrom-Json }
                    return $null
                }
            }

            $status = [int]$resp.StatusCode
            if ($status -eq 401 -and $attempt -lt $MaxRetries) {
                if ($bodyText -match 'InvalidAuthenticationToken|token is expired|Lifetime validation failed|Access token has expired') {
                    $newToken = Refresh-TokenFromProfile -Token $authToken
                    if ($newToken) {
                        Write-Log "WARN" "Token expire detecte. Nouveau token obtenu, retry $attempt."
                        $authToken = $newToken
                        Start-Sleep -Milliseconds 200
                        continue
                    }
                }
            }
            if (($status -eq 429 -or $status -ge 500) -and $attempt -lt $MaxRetries) {
                $delay = [int]([math]::Pow(2, $attempt) * 500)
                Write-Log "WARN" "Retry $attempt (HTTP $status). Attente ${delay}ms"
                Start-Sleep -Milliseconds $delay
                continue
            }
            throw "Graph HTTP $status : $bodyText"
        } catch {
            if ($attempt -lt $MaxRetries -and $_.Exception.Message -notmatch 'Graph HTTP') {
                Start-Sleep -Milliseconds ([int]([math]::Pow(2, $attempt) * 500))
                continue
            }
            throw
        }
    }
}

# =========================
# APPEL SHAREPOINT REST AVEC RETRY + REFRESH TOKEN
# =========================
function Invoke-SpRestCall {
    param(
        [string]$Method = "GET",
        [string]$Uri,
        [string]$Token,
        [string]$Body = "",
        [string]$ContentType = "application/json",
        [hashtable]$Headers = $null,
        [int]$MaxRetries = 4
    )

    $authToken = Resolve-ActiveToken -Token $Token
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::new($Method), $Uri)
        $req.Headers.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $authToken)
        $req.Headers.TryAddWithoutValidation("Accept", "application/json;odata=nometadata") | Out-Null

        if ($Headers) {
            foreach ($k in $Headers.Keys) {
                $req.Headers.TryAddWithoutValidation([string]$k, [string]$Headers[$k]) | Out-Null
            }
        }

        if ($Body) {
            $req.Content = [System.Net.Http.StringContent]::new($Body, [System.Text.Encoding]::UTF8, $ContentType)
        }

        $resp = $script:HttpClient.SendAsync($req).Result
        $respBody = $resp.Content.ReadAsStringAsync().Result
        $status = [int]$resp.StatusCode

        if ($resp.IsSuccessStatusCode) {
            $json = $null
            if ($respBody) {
                try { $json = $respBody | ConvertFrom-Json } catch {}
            }
            return @{ Status = $status; Body = $respBody; Json = $json; Token = $authToken }
        }

        if ($status -eq 401 -and $attempt -lt $MaxRetries) {
            if ($respBody -match 'InvalidAuthenticationToken|token is expired|Lifetime validation failed|Access token has expired') {
                $newToken = Refresh-TokenFromProfile -Token $authToken
                if ($newToken) {
                    Write-Log "WARN" "Token SharePoint expire detecte. Nouveau token obtenu, retry $attempt."
                    $authToken = $newToken
                    Start-Sleep -Milliseconds 200
                    continue
                }
            }
        }

        if (($status -eq 429 -or $status -ge 500) -and $attempt -lt $MaxRetries) {
            $delay = [int]([math]::Pow(2, $attempt) * 500)
            Write-Log "WARN" "Retry SharePoint $attempt (HTTP $status). Attente ${delay}ms"
            Start-Sleep -Milliseconds $delay
            continue
        }

        throw "SP REST $status : $respBody"
    }
}

# =========================
# RÉCUPÉRATION DU DRIVE ID
# =========================
function Get-DriveId {
    param(
        [string]$UPN,
        [string]$UserFolder,
        [string]$Token,
        [string]$Label
    )
    Write-Log "INFO" "$Label - UPN : $UPN"

    # Méthode 1 : direct par UPN
    try {
        $encUpn   = [Uri]::EscapeDataString($UPN)
        $driveUri = "https://graph.microsoft.com/v1.0/users/" + $encUpn + "/drive"
        $drive    = Invoke-GraphCall -Uri $driveUri -Token $Token
        Write-Log "SUCCESS" "$Label - Drive trouve via UPN : $($drive.id)"
        return $drive.id
    } catch {
        Write-Log "WARN" "$Label - Methode UPN directe echouee ($($_.Exception.Message)), tentative par liste..."
    }

    # Méthode 2 : chercher par userPrincipalName
    try {
        $userPart  = ($UPN -split '@')[0]
        $filterStr = "startsWith(userPrincipalName,'" + $userPart + "')"
        $searchUri = "https://graph.microsoft.com/v1.0/users?" + '$filter=' + [Uri]::EscapeDataString($filterStr) + "&" + '$select=id,userPrincipalName,displayName'
        $result    = Invoke-GraphCall -Uri $searchUri -Token $Token

        if ($result.value -and $result.value.Count -gt 0) {
            $user = $result.value[0]
            Write-Log "INFO" "$Label - Utilisateur trouve : $($user.userPrincipalName)"
            $driveUri = "https://graph.microsoft.com/v1.0/users/" + $user.id + "/drive"
            $drive    = Invoke-GraphCall -Uri $driveUri -Token $Token
            Write-Log "SUCCESS" "$Label - Drive trouve via recherche : $($drive.id)"
            return $drive.id
        }
        Write-Log "WARN" "$Label - Aucun utilisateur trouve avec le filtre '$filterStr'"
    } catch {
        Write-Log "WARN" "$Label - Methode recherche echouee : $($_.Exception.Message)"
    }

    throw "Impossible de trouver le drive pour '$UPN' (folder: $UserFolder). Verifiez l'UPN."
}

# =========================
# INFOS SITE SHAREPOINT
# =========================
function Get-DriveSiteInfo {
    param([string]$DriveId, [string]$Token)
    $drive   = Invoke-GraphCall -Uri ("https://graph.microsoft.com/v1.0/drives/" + $DriveId + "?`$select=webUrl,sharepointIds") -Token $Token
    $uri     = [Uri]$drive.webUrl
    $hostUrl = $uri.Scheme + "://" + $uri.Host
    $libPath = $uri.AbsolutePath
    $parts   = $libPath.TrimEnd('/') -split '/'
    $sitePath = ($parts[0..($parts.Count - 2)]) -join '/'

    # Resolve the Graph composite siteId ({hostname},{siteGuid},{webGuid}) via the site URL
    $graphSiteId = $null
    try {
        $siteRef = $uri.Host + ":" + $sitePath + ":"
        $siteObj = Invoke-GraphCall -Uri ("https://graph.microsoft.com/v1.0/sites/" + $siteRef) -Token $Token
        $graphSiteId = $siteObj.id
    } catch {}
    if (-not $graphSiteId -and $drive.sharepointIds -and $drive.sharepointIds.siteId) {
        $graphSiteId = $drive.sharepointIds.siteId
    }
    return @{ HostUrl = $hostUrl; SitePath = $sitePath; LibraryPath = $libPath; SiteId = $graphSiteId }
}

# Resolve a user email to their SharePoint site user lookupId
function Get-SharePointUserLookupId {
    param([string]$SiteId, [string]$Email, [string]$Token)
    $encFilter = [Uri]::EscapeDataString("fields/EMail eq '$Email'")
    # 'UserInfo' is the internal list name; 'User Information List' is the display name fallback
    foreach ($listName in @("UserInfo", "User%20Information%20List")) {
        try {
            $uri    = "https://graph.microsoft.com/v1.0/sites/" + $SiteId + "/lists/" + $listName + "/items?`$filter=" + $encFilter + "&`$select=id"
            $result = Invoke-GraphCall -Uri $uri -Token $Token
            if ($result -and $result.value -and $result.value.Count -gt 0) { return [string]$result.value[0].id }
        } catch {}
    }
    return $null
}

function Normalize-Email {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $v = $Value.Trim().ToLower()
    if ($v -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { return $null }
    return $v
}

function Resolve-MappedEmail {
    param(
        [string]$SourceEmail,
        [string]$FallbackTargetEmail = ""
    )

    $src = Normalize-Email -Value $SourceEmail
    if ($src -and $script:UserMapping.ContainsKey($src)) {
        $mapped = Normalize-Email -Value $script:UserMapping[$src]
        if ($mapped) { return $mapped }
    }
    if ($src) { return $src }

    $fallback = Normalize-Email -Value $FallbackTargetEmail
    if ($fallback) { return $fallback }
    return $null
}

function Get-ServerRelativePathForItem {
    param(
        [string]$LibraryPath,
        [string]$ItemRelPath
    )

    $lib = ($LibraryPath -replace '\\', '/').TrimEnd('/')
    $rel = ($ItemRelPath -replace '\\', '/').Trim('/')
    if ([string]::IsNullOrWhiteSpace($rel)) { return $lib }

    $libLeaf = [System.IO.Path]::GetFileName($lib)
    $parts = $rel -split '/'
    if ($parts.Count -gt 0 -and $libLeaf -and $parts[0].ToLower() -eq $libLeaf.ToLower()) {
        $rel = (($parts | Select-Object -Skip 1) -join '/')
    }

    if ([string]::IsNullOrWhiteSpace($rel)) { return $lib }
    return ($lib + "/" + $rel)
}

function Get-ListItemAllFieldsApiUrl {
    param(
        [string]$HostUrl,
        [string]$SitePath,
        [string]$ServerRelPath,
        [switch]$UseFolderApi
    )

    $safePath = $ServerRelPath -replace "'", "''"
    if ($UseFolderApi) {
        return $HostUrl + $SitePath + "/_api/web/GetFolderByServerRelativePath(decodedurl='" + $safePath + "')/ListItemAllFields"
    }
    return $HostUrl + $SitePath + "/_api/web/GetFileByServerRelativePath(decodedurl='" + $safePath + "')/ListItemAllFields"
}

function Get-ListItemAllFieldsApiUrlCandidates {
    param(
        [string]$HostUrl,
        [string]$SitePath,
        [string]$ServerRelPath,
        [switch]$UseFolderApi
    )

    $safePath = $ServerRelPath -replace "'", "''"
    if ($UseFolderApi) {
        return @(
            $HostUrl + $SitePath + "/_api/web/GetFolderByServerRelativePath(decodedurl='" + $safePath + "')/ListItemAllFields",
            $HostUrl + $SitePath + "/_api/web/GetFolderByServerRelativeUrl('" + $safePath + "')/ListItemAllFields"
        )
    }
    return @(
        $HostUrl + $SitePath + "/_api/web/GetFileByServerRelativePath(decodedurl='" + $safePath + "')/ListItemAllFields",
        $HostUrl + $SitePath + "/_api/web/GetFileByServerRelativeUrl('" + $safePath + "')/ListItemAllFields"
    )
}

function Get-SharePointListItemApiUrlCandidates {
    param(
        [string]$HostUrl,
        [string]$SitePath,
        [string]$ListTitle,
        [string]$ListItemId
    )

    if ([string]::IsNullOrWhiteSpace($ListTitle) -or [string]::IsNullOrWhiteSpace($ListItemId)) { return @() }
    $safeTitle = $ListTitle.Replace("'", "''")
    return @(
        $HostUrl + $SitePath + "/_api/web/lists/getByTitle('" + $safeTitle + "')/items(" + [int]$ListItemId + ")"
    )
}

function Get-DriveItemListItemId {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$Token
    )

    try {
        $uri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/items/" + $ItemId + "/listItem?`$select=id"
        $res = Invoke-GraphCall -Uri $uri -Token $Token
        if ($res -and $res.id) { return [string]$res.id }
    } catch {}
    return $null
}

function Resolve-SharePointUserEmailById {
    param(
        [string]$SpToken,
        [string]$HostUrl,
        [string]$SitePath,
        [string]$LookupId
    )

    if ([string]::IsNullOrWhiteSpace($LookupId)) { return $null }
    if (-not $script:SpUserEmailCache) { $script:SpUserEmailCache = @{} }

    $cacheKey = ($HostUrl + $SitePath + "|" + $LookupId).ToLower()
    if ($script:SpUserEmailCache.ContainsKey($cacheKey)) {
        return [string]$script:SpUserEmailCache[$cacheKey]
    }

    try {
        $uri = $HostUrl + $SitePath + "/_api/web/siteusers/getbyid(" + [int]$LookupId + ")?`$select=Email,UserPrincipalName,LoginName"
        $res = Invoke-SpRestCall -Method "GET" -Uri $uri -Token $SpToken
        if ($res.Token) { $script:DstSpToken = [string]$res.Token }
        $j = $res.Json

        $email = $null
        try { if ($j.Email) { $email = Normalize-Email -Value ([string]$j.Email) } } catch {}
        if (-not $email) { try { if ($j.UserPrincipalName) { $email = Normalize-Email -Value ([string]$j.UserPrincipalName) } } catch {} }
        if (-not $email) {
            try {
                if ($j.LoginName -and ([string]$j.LoginName).Contains('|')) {
                    $candidate = ([string]$j.LoginName).Split('|')[-1]
                    $email = Normalize-Email -Value $candidate
                }
            } catch {}
        }

        if ($email) { $script:SpUserEmailCache[$cacheKey] = $email }
        return $email
    } catch {
        return $null
    }
}

# Set Editor and optionally Modified date in a single ValidateUpdateListItem call
function Set-SharePointEditor {
    param(
        [string]$SpToken,
        [string]$HostUrl,
        [string]$SitePath,
        [string]$ServerRelPath,
        [string]$EditorEmail,
        [string]$ModifiedDate = ""  # ISO 8601 UTC; when set, also updates Modified in the same call
    )
    $claim = "i:0#.f|membership|$EditorEmail"
    $formValues = @(
        @{
            FieldName  = "Editor"
            FieldValue = ('[{"Key":"' + $claim + '"}]')
        }
    )
    if ($ModifiedDate) {
        $parsed = $null
        try { $parsed = [datetime]$ModifiedDate } catch {}
        if ($parsed) {
            # VVI expects a locale date string more reliably than ISO 8601 in many tenants.
            $formValues += @{
                FieldName  = "Modified"
                FieldValue = $parsed.ToString("dd/MM/yyyy HH:mm:ss")
            }
        }
    }
    $body = @{
        formValues         = $formValues
        bNewDocumentUpdate = $true
        checkInComment     = ""
    } | ConvertTo-Json -Depth 10 -Compress
    $baseCandidates = Get-ListItemAllFieldsApiUrlCandidates -HostUrl $HostUrl -SitePath $SitePath -ServerRelPath $ServerRelPath
    $res = $null
    $lastErr = $null
    foreach ($base in $baseCandidates) {
        $apiUrl = $base + "/ValidateUpdateListItem()"
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            try {
                $res = Invoke-SpRestCall -Method "POST" -Uri $apiUrl -Token $SpToken -Body $body
                break
            } catch {
                $lastErr = $_
                if ($_.Exception.Message -match 'SP REST 404' -and $attempt -lt 3) {
                    Start-Sleep -Milliseconds (300 * $attempt)
                    continue
                }
                break
            }
        }
        if ($res) { break }
    }
    if (-not $res) {
        if ($lastErr) { throw $lastErr }
        throw "Echec mise a jour Editor sur '$ServerRelPath'"
    }
    if ($res.Token) { $script:DstSpToken = [string]$res.Token }
    $respBody = [string]$res.Body
    # Log VVI response to show what fields were actually updated
    try {
        $vviResult = $respBody | ConvertFrom-Json
        $editorErr = ($vviResult.value | Where-Object { $_.FieldName -eq 'Editor' }).ErrorMessage
        $modErr    = ($vviResult.value | Where-Object { $_.FieldName -eq 'Modified' }).ErrorMessage
        Write-Log "DEBUG" "VVI reponse : Editor=$editorErr | Modified=$modErr"
    } catch {}
    # Read back the actual Editor after VVI to verify
    try {
        $people = Get-SharePointItemPeople -SpToken ([string]$res.Token) -HostUrl $HostUrl -SitePath $SitePath -ServerRelPath $ServerRelPath
        Write-Log "DEBUG" "VVI verification : Editor=$($people.Editor) / Modified=$($people.Modified)"
        return (Normalize-Email -Value $people.Editor)
    } catch { Write-Log "DEBUG" "VVI verification echouee : $($_.Exception.Message)" }
    return $null
}

# Provision user in the site and return their site user Id (EnsureUser — requires Sites.Manage.All on SharePoint)
function Invoke-SharePointEnsureUser {
    param([string]$SpToken, [string]$HostUrl, [string]$SitePath, [string]$Email)

    $emailNorm = Normalize-Email -Value $Email
    if (-not $emailNorm) { return $null }
    $cacheKey = ($HostUrl + $SitePath + "|" + $emailNorm).ToLower()
    if ($script:EnsureUserIdCache.ContainsKey($cacheKey)) {
        return [string]$script:EnsureUserIdCache[$cacheKey]
    }

    $body   = '{"logonName":"i:0#.f|membership|' + $Email + '"}'
    $apiUrl = $HostUrl + $SitePath + "/_api/web/EnsureUser"
    $res = Invoke-SpRestCall -Method "POST" -Uri $apiUrl -Token $SpToken -Body $body
    if ($res.Token) { $script:DstSpToken = [string]$res.Token }
    $json = $res.Json
    if (-not $json) { return $null }
    $id = [string]$json.Id
    if ($id) { $script:EnsureUserIdCache[$cacheKey] = $id }
    return $id
}

# Fallback when VVI claim update does not stick: set EditorId directly on ListItemAllFields.
function Set-SharePointEditorByLookupId {
    param(
        [string]$SpToken,
        [string]$HostUrl,
        [string]$SitePath,
        [string]$ServerRelPath,
        [string]$EditorLookupId
    )

    if ([string]::IsNullOrWhiteSpace($EditorLookupId)) { return $null }
    $body     = ('{"EditorId":' + [int]$EditorLookupId + '}')

    $baseCandidates = Get-ListItemAllFieldsApiUrlCandidates -HostUrl $HostUrl -SitePath $SitePath -ServerRelPath $ServerRelPath
    $res = $null
    $lastErr = $null
    foreach ($apiUrl in $baseCandidates) {
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            try {
                $res = Invoke-SpRestCall -Method "POST" -Uri $apiUrl -Token $SpToken -Body $body -Headers @{ "IF-MATCH" = "*"; "X-HTTP-Method" = "MERGE" }
                break
            } catch {
                $lastErr = $_
                if ($_.Exception.Message -match 'SP REST 404' -and $attempt -lt 3) {
                    Start-Sleep -Milliseconds (300 * $attempt)
                    continue
                }
                break
            }
        }
        if ($res) { break }
    }
    if (-not $res) {
        if ($lastErr) { throw $lastErr }
        throw "Echec MERGE EditorId sur '$ServerRelPath'"
    }
    if ($res.Token) { $script:DstSpToken = [string]$res.Token }

    try {
        $people = Get-SharePointItemPeople -SpToken ([string]$res.Token) -HostUrl $HostUrl -SitePath $SitePath -ServerRelPath $ServerRelPath
        Write-Log "DEBUG" "MERGE EditorId verification : Editor=$($people.Editor)"
        return (Normalize-Email -Value $people.Editor)
    } catch {
        Write-Log "DEBUG" "MERGE EditorId verification echouee : $($_.Exception.Message)"
    }

    return $null
}

function Get-SharePointItemPeople {
    param(
        [string]$SpToken,
        [string]$HostUrl,
        [string]$SitePath,
        [string]$ServerRelPath,
        [switch]$IsFolder,
        [string]$ListTitle,
        [string]$ListItemId
    )

    if ($ListTitle -and $ListItemId) {
        $baseCandidates = Get-SharePointListItemApiUrlCandidates -HostUrl $HostUrl -SitePath $SitePath -ListTitle $ListTitle -ListItemId $ListItemId
    } else {
        $baseCandidates = Get-ListItemAllFieldsApiUrlCandidates -HostUrl $HostUrl -SitePath $SitePath -ServerRelPath $ServerRelPath -UseFolderApi:$IsFolder
    }
    if (-not $IsFolder -and -not ($ListTitle -and $ListItemId)) {
        $folderCandidates = Get-ListItemAllFieldsApiUrlCandidates -HostUrl $HostUrl -SitePath $SitePath -ServerRelPath $ServerRelPath -UseFolderApi
        foreach ($u in $folderCandidates) { if ($baseCandidates -notcontains $u) { $baseCandidates += $u } }
    }

    $getRes = $null
    $lastErr = $null
    foreach ($baseUrl in $baseCandidates) {
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            try {
                $getRes = Invoke-SpRestCall -Method "GET" -Uri ($baseUrl + "?`$select=Author/EMail,Editor/EMail,Created,Modified&`$expand=Author,Editor") -Token $SpToken
                break
            } catch {
                $lastErr = $_
                $msg = $_.Exception.Message
                if ($msg -match 'SP REST 400' -and $msg -match '\$expand') {
                    try {
                        $getRes = Invoke-SpRestCall -Method "GET" -Uri ($baseUrl + "?`$select=AuthorId,EditorId,Created,Modified") -Token $SpToken
                        break
                    } catch {
                        $lastErr = $_
                    }
                }
                if ($msg -match 'SP REST 404' -and $attempt -lt 3) {
                    Start-Sleep -Milliseconds (300 * $attempt)
                    continue
                }
                break
            }
        }
        if ($getRes) { break }
    }
    if (-not $getRes) {
        if ($lastErr) { throw $lastErr }
        throw "Impossible de lire Author/Editor pour '$ServerRelPath'"
    }
    if ($getRes.Token) { $script:DstSpToken = [string]$getRes.Token }
    $j = $getRes.Json

    $author = $null
    $editor = $null
    try { $author = Normalize-Email -Value ([string]$j.Author.EMail) } catch {}
    try { $editor = Normalize-Email -Value ([string]$j.Editor.EMail) } catch {}

    if (-not $author) {
        $authorId = $null
        try { if ($j.AuthorId) { $authorId = [string]$j.AuthorId } } catch {}
        if ($authorId) { $author = Resolve-SharePointUserEmailById -SpToken ([string]$getRes.Token) -HostUrl $HostUrl -SitePath $SitePath -LookupId $authorId }
    }
    if (-not $editor) {
        $editorId = $null
        try { if ($j.EditorId) { $editorId = [string]$j.EditorId } } catch {}
        if ($editorId) { $editor = Resolve-SharePointUserEmailById -SpToken ([string]$getRes.Token) -HostUrl $HostUrl -SitePath $SitePath -LookupId $editorId }
    }

    return [pscustomobject]@{
        Author   = $author
        Editor   = $editor
        Created  = if ($j) { [string]$j.Created } else { "" }
        Modified = if ($j) { [string]$j.Modified } else { "" }
    }
}

function Set-SharePointAuthorEditor {
    param(
        [string]$SpToken,
        [string]$HostUrl,
        [string]$SitePath,
        [string]$ServerRelPath,
        [string]$AuthorEmail,
        [string]$EditorEmail,
        [string]$CreatedDate = "",
        [string]$ModifiedDate = "",
        [switch]$IsFolder,
        [switch]$SkipReadback,
        [string]$ListTitle,
        [string]$ListItemId
    )

    $formValues = @()

    $authorNorm = Normalize-Email -Value $AuthorEmail
    if ($authorNorm) {
        $claimAuthor = "i:0#.f|membership|$authorNorm"
        $formValues += @{
            FieldName  = "Author"
            FieldValue = ('[{"Key":"' + $claimAuthor + '"}]')
        }
    }

    $editorNorm = Normalize-Email -Value $EditorEmail
    if ($editorNorm) {
        $claimEditor = "i:0#.f|membership|$editorNorm"
        $formValues += @{
            FieldName  = "Editor"
            FieldValue = ('[{"Key":"' + $claimEditor + '"}]')
        }
    }

    if ($CreatedDate) {
        $createdParsed = $null
        try { $createdParsed = [datetime]$CreatedDate } catch {}
        if ($createdParsed) {
            $formValues += @{
                FieldName  = "Created"
                FieldValue = $createdParsed.ToString("dd/MM/yyyy HH:mm:ss")
            }
        }
    }

    if ($ModifiedDate) {
        $modifiedParsed = $null
        try { $modifiedParsed = [datetime]$ModifiedDate } catch {}
        if ($modifiedParsed) {
            $formValues += @{
                FieldName  = "Modified"
                FieldValue = $modifiedParsed.ToString("dd/MM/yyyy HH:mm:ss")
            }
        }
    }

    if ($formValues.Count -eq 0) { return $null }

    $body = @{
        formValues         = $formValues
        bNewDocumentUpdate = $true
        checkInComment     = ""
    } | ConvertTo-Json -Depth 10 -Compress

    if ($ListTitle -and $ListItemId) {
        $baseCandidates = Get-SharePointListItemApiUrlCandidates -HostUrl $HostUrl -SitePath $SitePath -ListTitle $ListTitle -ListItemId $ListItemId
    } else {
        $baseCandidates = Get-ListItemAllFieldsApiUrlCandidates -HostUrl $HostUrl -SitePath $SitePath -ServerRelPath $ServerRelPath -UseFolderApi:$IsFolder
    }
    if (-not $IsFolder -and -not ($ListTitle -and $ListItemId)) {
        $folderCandidates = Get-ListItemAllFieldsApiUrlCandidates -HostUrl $HostUrl -SitePath $SitePath -ServerRelPath $ServerRelPath -UseFolderApi
        foreach ($u in $folderCandidates) { if ($baseCandidates -notcontains $u) { $baseCandidates += $u } }
    }

    $res = $null
    $lastErr = $null
    foreach ($base in $baseCandidates) {
        $apiUrl = $base + "/ValidateUpdateListItem()"
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            try {
                $res = Invoke-SpRestCall -Method "POST" -Uri $apiUrl -Token $SpToken -Body $body
                if ($base -like "*GetFolderBy*") { $IsFolder = $true }
                break
            } catch {
                $lastErr = $_
                if ($_.Exception.Message -match 'SP REST 404' -and $attempt -lt 3) {
                    Start-Sleep -Milliseconds (300 * $attempt)
                    continue
                }
                break
            }
        }
        if ($res) { break }
    }
    if (-not $res) {
        if ($lastErr) { throw $lastErr }
        throw "Echec restauration Author/Editor sur '$ServerRelPath'"
    }
    if ($res.Token) { $script:DstSpToken = [string]$res.Token }

    if ($SkipReadback) {
        return [pscustomobject]@{
            Author   = (Normalize-Email -Value $authorNorm)
            Editor   = (Normalize-Email -Value $editorNorm)
            Created  = $CreatedDate
            Modified = $ModifiedDate
        }
    }

    return Get-SharePointItemPeople -SpToken ([string]$res.Token) -HostUrl $HostUrl -SitePath $SitePath -ServerRelPath $ServerRelPath -IsFolder:$IsFolder -ListTitle $ListTitle -ListItemId $ListItemId
}

function Set-SharePointAuthorEditorByLookupId {
    param(
        [string]$SpToken,
        [string]$HostUrl,
        [string]$SitePath,
        [string]$ServerRelPath,
        [string]$AuthorLookupId,
        [string]$EditorLookupId,
        [switch]$IsFolder,
        [switch]$SkipReadback,
        [string]$ListTitle,
        [string]$ListItemId
    )

    $payload = @{}
    if ($AuthorLookupId) { $payload.AuthorId = [int]$AuthorLookupId }
    if ($EditorLookupId) { $payload.EditorId = [int]$EditorLookupId }
    if ($payload.Count -eq 0) { return $null }

    $body = $payload | ConvertTo-Json -Compress
    if ($ListTitle -and $ListItemId) {
        $baseCandidates = Get-SharePointListItemApiUrlCandidates -HostUrl $HostUrl -SitePath $SitePath -ListTitle $ListTitle -ListItemId $ListItemId
    } else {
        $baseCandidates = Get-ListItemAllFieldsApiUrlCandidates -HostUrl $HostUrl -SitePath $SitePath -ServerRelPath $ServerRelPath -UseFolderApi:$IsFolder
    }
    if (-not $IsFolder -and -not ($ListTitle -and $ListItemId)) {
        $folderCandidates = Get-ListItemAllFieldsApiUrlCandidates -HostUrl $HostUrl -SitePath $SitePath -ServerRelPath $ServerRelPath -UseFolderApi
        foreach ($u in $folderCandidates) { if ($baseCandidates -notcontains $u) { $baseCandidates += $u } }
    }

    $res = $null
    $lastErr = $null
    foreach ($apiUrl in $baseCandidates) {
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            try {
                $res = Invoke-SpRestCall -Method "POST" -Uri $apiUrl -Token $SpToken -Body $body -Headers @{ "IF-MATCH" = "*"; "X-HTTP-Method" = "MERGE" }
                if ($apiUrl -like "*GetFolderBy*") { $IsFolder = $true }
                break
            } catch {
                $lastErr = $_
                if ($_.Exception.Message -match 'SP REST 404' -and $attempt -lt 3) {
                    Start-Sleep -Milliseconds (300 * $attempt)
                    continue
                }
                break
            }
        }
        if ($res) { break }
    }
    if (-not $res) {
        if ($lastErr) { throw $lastErr }
        throw "Echec MERGE AuthorId/EditorId sur '$ServerRelPath'"
    }
    if ($res.Token) { $script:DstSpToken = [string]$res.Token }

    if ($SkipReadback) {
        $authorEmail = $null
        $editorEmail = $null
        if ($AuthorLookupId) { $authorEmail = Resolve-SharePointUserEmailById -SpToken ([string]$res.Token) -HostUrl $HostUrl -SitePath $SitePath -LookupId $AuthorLookupId }
        if ($EditorLookupId) { $editorEmail = Resolve-SharePointUserEmailById -SpToken ([string]$res.Token) -HostUrl $HostUrl -SitePath $SitePath -LookupId $EditorLookupId }
        return [pscustomobject]@{
            Author   = $authorEmail
            Editor   = $editorEmail
            Created  = ""
            Modified = ""
        }
    }

    return Get-SharePointItemPeople -SpToken ([string]$res.Token) -HostUrl $HostUrl -SitePath $SitePath -ServerRelPath $ServerRelPath -IsFolder:$IsFolder -ListTitle $ListTitle -ListItemId $ListItemId
}

# =========================
# SCAN RÉCURSIF DU DRIVE
# =========================
function Get-AllDriveItems {
    param(
        [string]$DriveId,
        [string]$ItemId      = "root",
        [string]$CurrentPath = "",
        [string]$Token
    )
    $allItems = @()
    $url      = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/items/" + $ItemId + "/children?" + '$top=200&$select=id,name,size,folder,file'

    do {
        try {
            $response = Invoke-GraphCall -Uri $url -Token $Token
            foreach ($item in $response.value) {
                $itemPath = if ($CurrentPath) { "$CurrentPath/$($item.name)" } else { $item.name }
                if ($item.folder) {
                    Write-Log "DEBUG" "   📂 $itemPath"
                    $allItems += [pscustomobject]@{
                        Name = $item.name; RelPath = $itemPath
                        Id = $item.id; IsFolder = $true; Size = 0
                    }
                    $allItems += Get-AllDriveItems -DriveId $DriveId -ItemId $item.id `
                                                   -CurrentPath $itemPath -Token $Token
                } else {
                    Write-Log "DEBUG" "   📄 $itemPath ($($item.size) o)"
                    $allItems += [pscustomobject]@{
                        Name = $item.name; RelPath = $itemPath
                        Id = $item.id; IsFolder = $false; Size = [long]$item.size
                    }
                }
            }
            $url = $response.'@odata.nextLink'
        } catch {
            Write-Log "ERROR" "Erreur scan '$CurrentPath' : $($_.Exception.Message)"
            $url = $null
        }
    } while ($url)

    return $allItems
}

# =========================
# TÉLÉCHARGEMENT
# =========================
function Download-DriveFile {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$Token,
        [int]$MaxRetries = 5
    )
    $uri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/items/" + $ItemId + "/content"
    return Invoke-GraphCall -Uri $uri -Token $Token -ReturnBytes -MaxRetries $MaxRetries
}

function Get-DriveItemMeta {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$Token
    )

    $uri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/items/" + $ItemId + "?" + '$select=id,name,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,fileSystemInfo'
    return Invoke-GraphCall -Uri $uri -Token $Token
}

function Get-DriveFileVersions {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$Token
    )

    $versions = New-Object System.Collections.Generic.List[object]
    $url = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/items/" + $ItemId + "/versions?" + '$top=200&$select=id,lastModifiedDateTime,lastModifiedBy,size'

    do {
        $resp = Invoke-GraphCall -Uri $url -Token $Token
        if ($resp -and $resp.value) {
            foreach ($v in $resp.value) {
                $versions.Add($v) | Out-Null
            }
        }
        $url = $resp.'@odata.nextLink'
    } while ($url)

    $ordered = @($versions | Sort-Object -Property {
        try { [datetime]$_.lastModifiedDateTime } catch { [datetime]::MinValue }
    })

    if ($MaxVersionsPerFile -gt 0 -and $ordered.Count -gt $MaxVersionsPerFile) {
        $ordered = @($ordered | Select-Object -Last $MaxVersionsPerFile)
    }

    return $ordered
}

function Test-NeedsVersionHistoryRepair {
    param(
        [object[]]$SrcVersions,
        [object[]]$DstVersions
    )

    if (-not $SrcVersions -or $SrcVersions.Count -eq 0) { return $false }
    if (-not $DstVersions -or $DstVersions.Count -eq 0) { return $true }
    if ($DstVersions.Count -ne $SrcVersions.Count) { return $true }

    foreach ($v in $DstVersions) {
        $userEmail = $null
        $appName = $null
        try { if ($v.lastModifiedBy.user.email) { $userEmail = Normalize-Email -Value ([string]$v.lastModifiedBy.user.email) } } catch {}
        try { if ($v.lastModifiedBy.application.displayName) { $appName = [string]$v.lastModifiedBy.application.displayName } } catch {}

        if (-not $userEmail) { return $true }
        if ($appName -and $appName -match '(?i)sharepoint') { return $true }
    }

    return $false
}

function Initialize-TempTransferDir {
    if (-not $UseTempFileTransfer) { return }
    if (-not (Test-Path -LiteralPath $TempTransferDir)) {
        New-Item -ItemType Directory -Path $TempTransferDir -Force | Out-Null
    }
}

function New-TempTransferFilePath {
    param([string]$RelPath)

    Initialize-TempTransferDir
    $name = [System.IO.Path]::GetFileName($RelPath)
    if ([string]::IsNullOrWhiteSpace($name)) { $name = "item.bin" }
    $safe = ($name -replace '[\\/:*?"<>|]', '_')
    if ($safe.Length -gt 80) { $safe = $safe.Substring(0, 80) }
    return (Join-Path $TempTransferDir (([guid]::NewGuid().ToString("N")) + "_" + $safe))
}

function Download-GraphContentToFile {
    param(
        [string]$Uri,
        [string]$Token,
        [string]$OutFilePath,
        [int]$MaxRetries = 5
    )

    $authToken = Resolve-ActiveToken -Token $Token
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $Uri)
            $req.Headers.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $authToken)
            if ($attempt -gt 1) {
                # Force a fresh TCP connection on retries to avoid stale pooled sockets.
                $req.Headers.ConnectionClose = $true
            }

            $resp = $null
            try {
                $resp = $script:HttpClient.SendAsync($req, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).Result
                if (-not $resp) {
                    Write-Log "WARN" "Telechargement Graph: reponse HTTP nulle (tentative $attempt/$MaxRetries), retry connexion dediee."

                    $fallbackClient = [System.Net.Http.HttpClient]::new()
                    $fallbackClient.Timeout = [TimeSpan]::FromMinutes(30)
                    $fallbackReq = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $Uri)
                    $fallbackReq.Headers.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $authToken)
                    $fallbackReq.Headers.ConnectionClose = $true
                    try {
                        $resp = $fallbackClient.SendAsync($fallbackReq, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).Result
                    } finally {
                        $fallbackReq.Dispose()
                        $fallbackClient.Dispose()
                    }

                    if (-not $resp) {
                        throw "Reponse HTTP nulle lors du telechargement Graph"
                    }
                }

                $status = [int]$resp.StatusCode

                if ($resp.IsSuccessStatusCode) {
                    if (-not $resp.Content) {
                        throw "Graph HTTP $status : contenu de reponse absent"
                    }
                    $streamTask = $resp.Content.ReadAsStreamAsync()
                    if (-not $streamTask) {
                        throw "Graph HTTP $status : flux de contenu indisponible"
                    }
                    $stream = $streamTask.Result
                    if (-not $stream) {
                        throw "Graph HTTP $status : flux de contenu null"
                    }
                    $fs = [System.IO.File]::Open($OutFilePath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
                    try { $stream.CopyTo($fs) } finally { $fs.Dispose(); $stream.Dispose() }
                    return
                }

                $bodyText = if ($resp.Content) { $resp.Content.ReadAsStringAsync().Result } else { "" }
                if ($status -eq 401 -and $attempt -lt $MaxRetries) {
                    if ($bodyText -match 'InvalidAuthenticationToken|token is expired|Lifetime validation failed|Access token has expired') {
                        $newToken = Refresh-TokenFromProfile -Token $authToken
                        if ($newToken) {
                            $authToken = $newToken
                            Start-Sleep -Milliseconds 200
                            continue
                        }
                    }
                }
                if (($status -eq 429 -or $status -ge 500) -and $attempt -lt $MaxRetries) {
                    $delay = [int]([math]::Pow(2, $attempt) * 500)
                    Start-Sleep -Milliseconds $delay
                    continue
                }
                throw "Graph HTTP $status : $bodyText"
            } finally {
                if ($resp) { $resp.Dispose() }
            }
        } catch {
            if ($attempt -lt $MaxRetries -and $_.Exception.Message -notmatch 'Graph HTTP') {
                Start-Sleep -Milliseconds ([int]([math]::Pow(2, $attempt) * 500))
                continue
            }
            throw
        } finally {
            if ($req) { $req.Dispose() }
        }
    }
}

function Download-DriveFileToTempFile {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$Token,
        [string]$RelPath,
        [int]$MaxRetries = 5
    )

    $outFile = New-TempTransferFilePath -RelPath $RelPath
    $uri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/items/" + $ItemId + "/content"
    Download-GraphContentToFile -Uri $uri -Token $Token -OutFilePath $outFile -MaxRetries $MaxRetries
    return $outFile
}

function Download-DriveFileVersionToTempFile {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$VersionId,
        [string]$Token,
        [string]$RelPath,
        [int]$MaxRetries = 5
    )

    $outFile = New-TempTransferFilePath -RelPath $RelPath
    $uri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/items/" + $ItemId + "/versions/" + $VersionId + "/content"
    try {
        Download-GraphContentToFile -Uri $uri -Token $Token -OutFilePath $outFile -MaxRetries $MaxRetries
        return $outFile
    } catch {
        if ($_.Exception.Message -match 'You cannot get the content of the current version') {
            Download-GraphContentToFile -Uri ("https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/items/" + $ItemId + "/content") -Token $Token -OutFilePath $outFile -MaxRetries $MaxRetries
            return $outFile
        }
        throw
    }
}

function Upload-FileToDrivePath {
    param(
        [string]$DriveId,
        [string]$RelPath,
        [string]$LocalPath,
        [string]$Token,
        [int]$ChunkBytes = 10485760
    )

    if (-not (Test-Path -LiteralPath $LocalPath)) {
        throw "Fichier temporaire introuvable: $LocalPath"
    }

    $segs = ($RelPath -replace '\\', '/') -split '/' | Where-Object { $_ } | ForEach-Object { [Uri]::EscapeDataString($_) }
    $encPath = $segs -join '/'
    $createUri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/root:/" + $encPath + ":/createUploadSession"
    $session = Invoke-GraphCall -Uri $createUri -Token $Token -Method "POST" -StringBody '{"item":{"@microsoft.graph.conflictBehavior":"replace"}}'
    if (-not $session -or -not $session.uploadUrl) {
        throw "Session d'upload invalide pour '$RelPath'"
    }
    $uploadUrl = $session.uploadUrl

    $fi = Get-Item -LiteralPath $LocalPath
    $total = [long]$fi.Length
    if ($total -eq 0) {
        Upload-SmallFile -DriveId $DriveId -RelPath $RelPath -Bytes ([byte[]]@()) -Token $Token
        return
    }

    $buffer = New-Object byte[] $ChunkBytes
    $fs = [System.IO.File]::Open($LocalPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try {
        $pos = 0L
        $attempt = 0
        while ($pos -lt $total) {
            $read = $fs.Read($buffer, 0, $buffer.Length)
            if ($read -le 0) { break }

            $attempt++
            $from = $pos
            $to   = $pos + $read - 1
            $content = [System.Net.Http.ByteArrayContent]::new($buffer, 0, $read)
            $content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::new("application/octet-stream")
            $null = $content.Headers.TryAddWithoutValidation("Content-Range", "bytes $from-$to/$total")

            $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Put, $uploadUrl)
            $req.Content = $content
            $resp = $script:HttpClient.SendAsync($req).Result
            if (-not $resp) {
                throw "Reponse HTTP nulle pendant upload de '$RelPath'"
            }

            if ($resp.IsSuccessStatusCode) {
                $pos += $read
                continue
            }

            $status = [int]$resp.StatusCode
            $errBody = if ($resp.Content) { $resp.Content.ReadAsStringAsync().Result } else { "" }
            if ($status -eq 429 -or ($status -ge 500 -and $status -lt 600) -or ($status -eq 409 -and $errBody -match 'nameAlreadyExists|currently being uploaded')) {
                $delay = [int]([math]::Pow(2, $attempt) * 500)
                Start-Sleep -Milliseconds $delay
                $fs.Position = $from
                continue
            }
            throw "Upload HTTP $status : $errBody"
        }
    } finally {
        $fs.Dispose()
    }
}

function Copy-DriveFileContent {
    param(
        [string]$SrcDriveId,
        [string]$SrcItemId,
        [string]$DstDriveId,
        [string]$RelPath,
        [string]$SrcToken,
        [string]$DstToken,
        [string]$VersionId = "",
        [switch]$UseCurrentContentForVersion,
        [switch]$AllowItemIdRefresh = $true,
        [int]$ChunkBytes = 10485760
    )

    $tempFile = $null
    try {
        try {
            if ($UseTempFileTransfer) {
                if ($UseCurrentContentForVersion -or [string]::IsNullOrWhiteSpace($VersionId)) {
                    $tempFile = Download-DriveFileToTempFile -DriveId $SrcDriveId -ItemId $SrcItemId -Token $SrcToken -RelPath $RelPath
                } else {
                    $tempFile = Download-DriveFileVersionToTempFile -DriveId $SrcDriveId -ItemId $SrcItemId -VersionId $VersionId -Token $SrcToken -RelPath $RelPath
                }

                if (-not $tempFile -or -not (Test-Path -LiteralPath $tempFile)) {
                    throw "Contenu temporaire introuvable pour '$RelPath'"
                }

                $size = 0L
                try { $size = [long](Get-Item -LiteralPath $tempFile).Length } catch {}
                Upload-FileToDrivePath -DriveId $DstDriveId -RelPath $RelPath -LocalPath $tempFile -Token $DstToken -ChunkBytes $ChunkBytes
                return $size
            }

            if ($UseCurrentContentForVersion -or [string]::IsNullOrWhiteSpace($VersionId)) {
                $bytes = Download-DriveFile -DriveId $SrcDriveId -ItemId $SrcItemId -Token $SrcToken
            } else {
                $bytes = Download-DriveFileVersion -DriveId $SrcDriveId -ItemId $SrcItemId -VersionId $VersionId -Token $SrcToken
            }
            if ($null -eq $bytes) { throw "Contenu source vide pour '$RelPath'" }

            Upload-BytesToDrivePath -DriveId $DstDriveId -RelPath $RelPath -Bytes $bytes -Token $DstToken
            return [long]$bytes.Length
        } catch {
            $msg = $_.Exception.Message
            if ($AllowItemIdRefresh -and $msg -match 'itemNotFound') {
                try {
                    $srcByPath = Get-DriveItemByRelPath -DriveId $SrcDriveId -RelPath $RelPath -Token $SrcToken
                    if ($srcByPath -and $srcByPath.id -and ([string]$srcByPath.id -ne [string]$SrcItemId)) {
                        Write-Log "WARN" "ID source rafraichi apres itemNotFound : $RelPath"
                        return Copy-DriveFileContent -SrcDriveId $SrcDriveId -SrcItemId ([string]$srcByPath.id) -DstDriveId $DstDriveId -RelPath $RelPath -SrcToken $SrcToken -DstToken $DstToken -VersionId $VersionId -UseCurrentContentForVersion:$UseCurrentContentForVersion -AllowItemIdRefresh:$false -ChunkBytes $ChunkBytes
                    }
                } catch {}
            }
            throw
        }
    } finally {
        if ($tempFile -and (Test-Path -LiteralPath $tempFile)) {
            Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
}

function Download-DriveFileVersion {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$VersionId,
        [string]$Token,
        [int]$MaxRetries = 5
    )

    $uri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/items/" + $ItemId + "/versions/" + $VersionId + "/content"
    return Invoke-GraphCall -Uri $uri -Token $Token -ReturnBytes -MaxRetries $MaxRetries
}

function Get-DriveItemByRelPath {
    param(
        [string]$DriveId,
        [string]$RelPath,
        [string]$Token
    )

    $segs    = ($RelPath -replace '\\', '/') -split '/' | Where-Object { $_ } |
               ForEach-Object { [Uri]::EscapeDataString($_) }
    $encPath = $segs -join '/'
    $uri     = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/root:/" + $encPath + "?" + '$select=id,name,createdDateTime,lastModifiedDateTime,fileSystemInfo'
    return Invoke-GraphCall -Uri $uri -Token $Token
}

function Set-DriveItemFileTimestamps {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$Token,
        [datetime]$CreatedUtc,
        [datetime]$ModifiedUtc
    )

    $body = @{
        fileSystemInfo = @{
            createdDateTime      = $CreatedUtc.ToUniversalTime().ToString("o")
            lastModifiedDateTime = $ModifiedUtc.ToUniversalTime().ToString("o")
        }
    } | ConvertTo-Json -Depth 5

    $uri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/items/" + $ItemId
    Invoke-GraphCall -Uri $uri -Token $Token -Method "PATCH" -StringBody $body | Out-Null
}

function Get-DriveItemListFields {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$Token
    )

    $uri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/items/" + $ItemId + "/listItem/fields"
    return Invoke-GraphCall -Uri $uri -Token $Token
}

function Get-CopiableListFields {
    param([object]$FieldsObj)

    if (-not $FieldsObj) { return @{} }

    $blocked = @(
        'id','ID','ContentType','ContentTypeId','FileLeafRef','FileRef','FileDirRef',
        'Created','Modified','Author','Editor','AuthorLookupId','EditorLookupId',
        'Created_x0020_By','Modified_x0020_By','_UIVersionString','owshiddenversion',
        'GUID','UniqueId','ComplianceAssetId','AppAuthor','AppEditor',
        'AppAuthorLookupId','AppEditorLookupId','DocIcon',
        'SMTotalSize','SMTotalFileStreamSize','_ModerationStatus','_HasCopyDestinations',
        '_CopySource','_Level','_IsCurrentVersion','_CheckinComment','LinkFilename',
        'LinkFilenameNoMenu','LinkTitle','LinkTitleNoMenu','ItemChildCount','FolderChildCount'
    )

    $result = @{}
    foreach ($prop in $FieldsObj.PSObject.Properties) {
        $name  = [string]$prop.Name
        $value = $prop.Value

        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ($name.StartsWith('_')) { continue }
        if ($blocked -contains $name) { continue }
        if ($name -like 'OData_*') { continue }
        if ($name -like '*StringId') { continue }

        if ($null -ne $value) {
            if ($value -is [System.Collections.IDictionary] -or ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string]))) {
                # Keep only scalar metadata fields to avoid Graph serialization errors.
                continue
            }
            $result[$name] = $value
        }
    }

    return $result
}

function Set-DriveItemListFields {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [hashtable]$Fields,
        [string]$Token
    )

    if (-not $Fields -or $Fields.Count -eq 0) { return }
    $body = $Fields | ConvertTo-Json -Depth 8
    $uri  = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/items/" + $ItemId + "/listItem/fields"
    Invoke-GraphCall -Uri $uri -Token $Token -Method "PATCH" -StringBody $body | Out-Null
}

function Set-DriveItemListFieldsBestEffort {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [hashtable]$Fields,
        [string]$Token,
        [string]$ItemRelPath
    )

    if (-not $Fields -or $Fields.Count -eq 0) { return 0 }

    $work = @{}
    foreach ($k in $Fields.Keys) { $work[$k] = $Fields[$k] }

    $maxAttempts = [Math]::Min(20, [Math]::Max(1, $work.Count))
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        if ($work.Count -eq 0) { break }
        try {
            Set-DriveItemListFields -DriveId $DriveId -ItemId $ItemId -Fields $work -Token $Token
            return $work.Count
        } catch {
            $msg = $_.Exception.Message
            $fieldName = $null

            if ($msg -match "Field '([^']+)' is read-only") { $fieldName = $Matches[1] }

            if ($fieldName -and $work.ContainsKey($fieldName)) {
                Write-Log "DEBUG" "Champ metadonnee ignore (read-only) : $ItemRelPath / $fieldName"
                $work.Remove($fieldName)
                continue
            }

            throw
        }
    }

    if ($work.Count -eq 0) {
        Write-Log "DEBUG" "Tous les champs metadonnees restants sont non modifiables : $ItemRelPath"
        return 0
    }

    throw "Echec application metadonnees (champs restants: $($work.Keys -join ', '))"
}

function Copy-DriveItemMetadata {
    param(
        [string]$SrcDriveId,
        [string]$SrcItemId,
        [string]$DstDriveId,
        [string]$DstItemId,
        [string]$SrcToken,
        [string]$DstToken,
        [string]$ItemRelPath
    )

    if (-not $CopyItemMetadata) { return }
    if ($ItemRelPath -eq "/") {
        Write-Log "DEBUG" "Metadonnees racine ignorees (listItem indisponible via Graph)"
        return
    }
    if (-not $SrcItemId -or -not $DstItemId) {
        Write-Log "WARN" "Metadonnees ignorees (item id manquant) : $ItemRelPath"
        return
    }

    try {
        $srcFieldsRaw = Get-DriveItemListFields -DriveId $SrcDriveId -ItemId $SrcItemId -Token $SrcToken
        $fieldsToCopy = Get-CopiableListFields -FieldsObj $srcFieldsRaw
        if (-not $fieldsToCopy -or $fieldsToCopy.Count -eq 0) {
            Write-Log "DEBUG" "Aucune metadonnee copiable pour '$ItemRelPath'"
            return
        }

        $appliedCount = Set-DriveItemListFieldsBestEffort -DriveId $DstDriveId -ItemId $DstItemId -Fields $fieldsToCopy -Token $DstToken -ItemRelPath $ItemRelPath
        Write-Log "DEBUG" "Metadonnees appliquees : $ItemRelPath ($appliedCount champ(s))"
    } catch {
        Write-Log "WARN" "Metadonnees non appliquees pour '$ItemRelPath' : $($_.Exception.Message)"
    }
}

function Get-DriveRootItem {
    param(
        [string]$DriveId,
        [string]$Token
    )

    $uri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/root?" + '$select=id,name'
    return Invoke-GraphCall -Uri $uri -Token $Token
}

function Get-DriveItemPermissions {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$Token
    )

    $allPerms = @()
    $safeItemId = [Uri]::EscapeDataString($ItemId)
    $url = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/items/" + $safeItemId + "/permissions?" + '$top=200'

    do {
        $resp = Invoke-GraphCall -Uri $url -Token $Token
        $page = $null
        if ($resp) {
            if ($resp.PSObject.Properties.Name -contains 'value') { $page = $resp.value }
            elseif ($resp -is [System.Array]) { $page = $resp }
        }
        if ($page) {
            $allPerms += @($page)
        }
        $url = $resp.'@odata.nextLink'
    } while ($url)

    return $allPerms
}

function Grant-DriveItemPermission {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$DstEmail,
        [string]$Role,
        [string]$Token
    )

    $inviteUri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/items/" + $ItemId + "/invite"
    $body = @{
        recipients     = @(@{ email = $DstEmail })
        requireSignIn  = $true
        sendInvitation = $false
        roles          = @($Role)
    } | ConvertTo-Json -Depth 6

    Invoke-GraphCall -Uri $inviteUri -Token $Token -Method "POST" -StringBody $body | Out-Null
}

function Get-EmailsFromPermission {
    param([object]$Permission)

    $emails = New-Object System.Collections.Generic.List[string]

    try {
        if ($Permission.grantedToV2 -and $Permission.grantedToV2.user -and $Permission.grantedToV2.user.email) {
            [void]$emails.Add([string]$Permission.grantedToV2.user.email)
        }
    } catch {}
    try {
        if ($Permission.grantedTo -and $Permission.grantedTo.user -and $Permission.grantedTo.user.email) {
            [void]$emails.Add([string]$Permission.grantedTo.user.email)
        }
    } catch {}
    try {
        if ($Permission.grantedToIdentitiesV2) {
            foreach ($idObj in $Permission.grantedToIdentitiesV2) {
                if ($idObj.user -and $idObj.user.email) {
                    [void]$emails.Add([string]$idObj.user.email)
                }
            }
        }
    } catch {}
    try {
        if ($Permission.grantedToIdentities) {
            foreach ($idObj in $Permission.grantedToIdentities) {
                if ($idObj.user -and $idObj.user.email) {
                    [void]$emails.Add([string]$idObj.user.email)
                }
            }
        }
    } catch {}

    return @($emails | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim().ToLower() } | Select-Object -Unique)
}

function Copy-DriveItemPermissions {
    param(
        [string]$SrcDriveId,
        [string]$SrcItemId,
        [string]$DstDriveId,
        [string]$DstItemId,
        [string]$SrcToken,
        [string]$DstToken,
        [string]$ItemRelPath
    )

    if (-not $CopyPermissions) { return }
    if (-not $SrcItemId -or -not $DstItemId) {
        Write-Log "WARN" "Permissions ignorees (item id manquant) : $ItemRelPath"
        return
    }

    try {
        $permissions = @(Get-DriveItemPermissions -DriveId $SrcDriveId -ItemId $SrcItemId -Token $SrcToken)
    } catch {
        Write-Log "WARN" "Lecture des permissions source impossible pour '$ItemRelPath' : $($_.Exception.Message)"
        return
    }

    foreach ($perm in $permissions) {
        # Do not copy inherited entries at item level; they flow from parent permissions.
        if ($perm.inheritedFrom) { continue }
        if ($perm.link) {
            Write-Log "DEBUG" "Lien de partage ignore pour '$ItemRelPath' (permission id=$($perm.id))"
            continue
        }

        $role = if ($perm.roles -and ($perm.roles -contains 'write')) { 'write' } else { 'read' }
        $emails = @(Get-EmailsFromPermission -Permission $perm)

        if ($emails.Count -eq 0) {
            Write-Log "DEBUG" "Permission sans email exploitable ignoree : $ItemRelPath (id=$($perm.id))"
            continue
        }

        foreach ($srcEmail in $emails) {
            if ($srcEmail -match '^spo-grid-all-users/') { continue }

            $dstEmail = if ($script:UserMapping.ContainsKey($srcEmail)) { $script:UserMapping[$srcEmail] } else { $srcEmail }
            if ([string]::IsNullOrWhiteSpace($dstEmail)) { continue }

            try {
                Grant-DriveItemPermission -DriveId $DstDriveId -ItemId $DstItemId -DstEmail $dstEmail -Role $role -Token $DstToken
                Write-Log "DEBUG" "Permission $role appliquee : $ItemRelPath ($srcEmail -> $dstEmail)"
            } catch {
                Write-Log "WARN" "Permission non appliquee : $ItemRelPath ($srcEmail -> $dstEmail, role=$role) : $($_.Exception.Message)"
            }
        }
    }
}

function Upload-BytesToDrivePath {
    param(
        [string]$DriveId,
        [string]$RelPath,
        [byte[]]$Bytes,
        [string]$Token
    )

    $limitPetit = 4 * 1024 * 1024
    $limitGrand = $ChunkSizeMB * 1024 * 1024

    if ($null -eq $Bytes) {
        throw "Upload impossible : contenu binaire null pour '$RelPath'"
    }

    if ($Bytes.Length -le $limitPetit) {
        Upload-SmallFile -DriveId $DriveId -RelPath $RelPath -Bytes $Bytes -Token $Token
    } else {
        Upload-LargeFile -DriveId $DriveId -RelPath $RelPath -Bytes $Bytes -Token $Token -ChunkBytes $limitGrand
    }
}

# =========================
# CRÉER UN DOSSIER
# =========================
function Ensure-DriveFolder {
    param(
        [string]$DriveId,
        [string]$FolderRelPath,
        [string]$Token
    )
    $segs = ($FolderRelPath -replace '\\', '/') -split '/' | Where-Object { $_ }
    if ($segs.Count -eq 0) { return }

    $currentPath = ""
    foreach ($seg in $segs) {
        $currentPath = if ($currentPath) { "$currentPath/$seg" } else { $seg }
        $cacheKey = Get-FolderCacheKey -DriveId $DriveId -FolderRelPath $currentPath
        if ($script:FolderExistsCache.ContainsKey($cacheKey)) { continue }

        $encSegs     = ($currentPath -split '/') | ForEach-Object { [Uri]::EscapeDataString($_) }
        $encCurrent  = $encSegs -join '/'
        $checkUri    = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/root:/" + $encCurrent
        $exists      = $true

        try {
            Invoke-GraphCall -Uri $checkUri -Token $Token | Out-Null
            $script:FolderExistsCache[$cacheKey] = $true
        } catch { $exists = $false }

        if (-not $exists) {
            $parts = $currentPath -split '/'
            $parentSegs = if ($parts.Count -gt 1) { $parts[0..($parts.Count - 2)] } else { @() }
            if ($parentSegs.Count -gt 0) {
                $encParent = ($parentSegs | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/'
                $createUri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/root:/" + $encParent + ":/children"
            } else {
                $createUri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/root/children"
            }
            $bodyJson = '{"name":"' + ($seg -replace '"', '\"') + '","folder":{},"@microsoft.graph.conflictBehavior":"replace"}'
            try {
                Invoke-GraphCall -Uri $createUri -Token $Token -Method "POST" -StringBody $bodyJson | Out-Null
                $script:FolderExistsCache[$cacheKey] = $true
                Write-Log "SUCCESS" "📂 Dossier cree : $currentPath"
            } catch {
                Write-Log "DEBUG" "Dossier deja present : $currentPath"
            }
        }
    }
}

# =========================
# UPLOAD PETIT FICHIER (< 4MB)
# =========================
function Upload-SmallFile {
    param(
        [string]$DriveId,
        [string]$RelPath,
        [byte[]]$Bytes,
        [string]$Token
    )
    if ($null -eq $Bytes) {
        throw "Upload petit fichier impossible : contenu null pour '$RelPath'"
    }

    $segs      = ($RelPath -replace '\\', '/') -split '/' | Where-Object { $_ } |
                 ForEach-Object { [Uri]::EscapeDataString($_) }
    $encPath   = $segs -join '/'
    $total     = $Bytes.Length

    # Upload sessions require a valid Content-Range; for empty files use direct PUT /content.
    if ($total -eq 0) {
        $uri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/root:/" + $encPath + ":/content"
        Invoke-GraphCall -Uri $uri -Token $Token -Method "PUT" -ByteBody ([byte[]]@()) | Out-Null
        return
    }

    # Use session instead of PUT so eTag is not checked on subsequent version uploads
    $createUri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/root:/" + $encPath + ":/createUploadSession"
    $session   = Invoke-GraphCall -Uri $createUri -Token $Token -Method "POST" `
                     -StringBody '{"item":{"@microsoft.graph.conflictBehavior":"replace"}}'
    $uploadUrl = $session.uploadUrl
    $content   = [System.Net.Http.ByteArrayContent]::new($Bytes)
    $content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::new("application/octet-stream")
    $null      = $content.Headers.TryAddWithoutValidation("Content-Range", "bytes 0-$($total - 1)/$total")
    $req       = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Put, $uploadUrl)
    $req.Content = $content
    $resp      = $script:HttpClient.SendAsync($req).Result
    if (-not $resp.IsSuccessStatusCode) {
        $errBody = $resp.Content.ReadAsStringAsync().Result
        throw "Upload HTTP $([int]$resp.StatusCode) : $errBody"
    }
}

# =========================
# UPLOAD GRAND FICHIER (>= 4MB)
# =========================
function Upload-LargeFile {
    param(
        [string]$DriveId,
        [string]$RelPath,
        [byte[]]$Bytes,
        [string]$Token,
        [int]$ChunkBytes = 10485760
    )
    if ($null -eq $Bytes) {
        throw "Upload grand fichier impossible : contenu null pour '$RelPath'"
    }

    $segs      = ($RelPath -replace '\\', '/') -split '/' | Where-Object { $_ } |
                 ForEach-Object { [Uri]::EscapeDataString($_) }
    $encPath   = $segs -join '/'
    $createUri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/root:/" + $encPath + ":/createUploadSession"
    $bodyJson  = '{"item":{"@microsoft.graph.conflictBehavior":"replace"}}'

    $session   = Invoke-GraphCall -Uri $createUri -Token $Token -Method "POST" -StringBody $bodyJson
    $uploadUrl = $session.uploadUrl
    $total     = $Bytes.Length
    $pos       = 0
    $attempt   = 0

    while ($pos -lt $total) {
        $attempt++
        $chunk   = [Math]::Min($ChunkBytes, $total - $pos)
        $from    = $pos
        $to      = $pos + $chunk - 1

        $content = [System.Net.Http.ByteArrayContent]::new($Bytes, $from, $chunk)
        $content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::new("application/octet-stream")
        $null    = $content.Headers.TryAddWithoutValidation("Content-Range", "bytes $from-$to/$total")

        $req         = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Put, $uploadUrl)
        $req.Content = $content
        $resp        = $script:HttpClient.SendAsync($req).Result

        if ($resp.IsSuccessStatusCode) {
            $pos += $chunk
            Write-Log "DEBUG" "   Chunk : $pos / $total"
        } else {
            $status = [int]$resp.StatusCode
            if ($status -eq 429 -or ($status -ge 500 -and $status -lt 600)) {
                $delay = [int]([math]::Pow(2, $attempt) * 500)
                Write-Log "WARN" "Chunk retry ($status). Attente ${delay}ms"
                Start-Sleep -Milliseconds $delay
                continue
            }
            $errBody = $resp.Content.ReadAsStringAsync().Result
            throw "Upload HTTP $status : $errBody"
        }
    }
}

# =========================
# COPIE D'UN ÉLÉMENT
# =========================
function Copy-DriveItem {
    param(
        [pscustomobject]$Item,
        [string]$SrcDriveId,
        [string]$DstDriveId,
        [string]$SrcToken,
        [string]$DstToken
    )
    try {
        if ($Item.IsFolder) {
            if ($script:RepairVersionHistoryOnlyResolved) {
                return
            }
            Ensure-DriveFolder -DriveId $DstDriveId -FolderRelPath $Item.RelPath -Token $DstToken
            $dstFolderId = $null
            $srcFolderMeta = $null
            try { $srcFolderMeta = Get-DriveItemMeta -DriveId $SrcDriveId -ItemId $Item.Id -Token $SrcToken } catch {}
            try {
                $dstFolder = Get-DriveItemByRelPath -DriveId $DstDriveId -RelPath $Item.RelPath -Token $DstToken
                if ($dstFolder -and $dstFolder.id) { $dstFolderId = [string]$dstFolder.id }
            } catch {}

            $skipFolderAuthorRestore = Test-ShouldSkipAuthorRestore -ItemRelPath $Item.RelPath
            if ($RestoreItemAuthorEditor -and -not $skipFolderAuthorRestore -and $dstFolderId -and $srcFolderMeta -and $script:DstSpToken -and $script:DstSiteInfo) {
                try {
                    $srcAuthor = $null
                    $srcEditor = $null
                    try { if ($srcFolderMeta.createdBy.user.email) { $srcAuthor = [string]$srcFolderMeta.createdBy.user.email } } catch {}
                    try { if ($srcFolderMeta.lastModifiedBy.user.email) { $srcEditor = [string]$srcFolderMeta.lastModifiedBy.user.email } } catch {}
                    if (-not $srcEditor) { $srcEditor = $srcAuthor }

                    $mappedAuthor = Resolve-MappedEmail -SourceEmail $srcAuthor -FallbackTargetEmail $script:CurrentTargetUpn
                    $mappedEditor = Resolve-MappedEmail -SourceEmail $srcEditor -FallbackTargetEmail $script:CurrentTargetUpn
                    $createdWhen  = if ($srcFolderMeta.createdDateTime) { [string]$srcFolderMeta.createdDateTime } else { "" }
                    $modifiedWhen = if ($srcFolderMeta.lastModifiedDateTime) { [string]$srcFolderMeta.lastModifiedDateTime } else { "" }

                    if ($mappedAuthor -or $mappedEditor) {
                        $authorId = $null
                        $editorId = $null
                        if ($mappedAuthor) { try { $authorId = Invoke-SharePointEnsureUser -SpToken $script:DstSpToken -HostUrl $script:DstSiteInfo.HostUrl -SitePath $script:DstSiteInfo.SitePath -Email $mappedAuthor } catch {} }
                        if ($mappedEditor) { try { $editorId = Invoke-SharePointEnsureUser -SpToken $script:DstSpToken -HostUrl $script:DstSiteInfo.HostUrl -SitePath $script:DstSiteInfo.SitePath -Email $mappedEditor } catch {} }

                        $serverRelPath = Get-ServerRelativePathForItem -LibraryPath $script:DstSiteInfo.LibraryPath -ItemRelPath $Item.RelPath
                        $dstListItemId = $null
                        try { $dstListItemId = Get-DriveItemListItemId -DriveId $DstDriveId -ItemId $dstFolderId -Token $DstToken } catch {}
                        $peopleState = Set-SharePointAuthorEditor -SpToken $script:DstSpToken -HostUrl $script:DstSiteInfo.HostUrl -SitePath $script:DstSiteInfo.SitePath -ServerRelPath $serverRelPath -AuthorEmail $mappedAuthor -EditorEmail $mappedEditor -CreatedDate $createdWhen -ModifiedDate $modifiedWhen -IsFolder -SkipReadback -ListTitle ([System.IO.Path]::GetFileName($script:DstSiteInfo.LibraryPath)) -ListItemId $dstListItemId

                        $authorOk = $false
                        $editorOk = $false
                        try { $authorOk = ($peopleState.Author -eq (Normalize-Email -Value $mappedAuthor)) } catch {}
                        try { $editorOk = ($peopleState.Editor -eq (Normalize-Email -Value $mappedEditor)) } catch {}

                        if (-not $authorOk -or -not $editorOk) {
                            $mergeState = Set-SharePointAuthorEditorByLookupId -SpToken $script:DstSpToken -HostUrl $script:DstSiteInfo.HostUrl -SitePath $script:DstSiteInfo.SitePath -ServerRelPath $serverRelPath -AuthorLookupId $authorId -EditorLookupId $editorId -IsFolder -SkipReadback -ListTitle ([System.IO.Path]::GetFileName($script:DstSiteInfo.LibraryPath)) -ListItemId $dstListItemId
                            try { $authorOk = ($mergeState.Author -eq (Normalize-Email -Value $mappedAuthor)) } catch {}
                            try { $editorOk = ($mergeState.Editor -eq (Normalize-Email -Value $mappedEditor)) } catch {}
                        }

                        if ($authorOk -and $editorOk) {
                            Write-Log "DEBUG" "Auteur/Modifie par restaures (dossier) : $($Item.RelPath)"
                        } else {
                            Write-AuthorRestoreIssue -ItemRelPath $Item.RelPath -Text "Auteur/Modifie par non verifies (dossier) : $($Item.RelPath)"
                        }
                    }
                } catch {
                    Write-AuthorRestoreIssue -ItemRelPath $Item.RelPath -Text "Restauration auteur/createur dossier echouee '$($Item.RelPath)' : $($_.Exception.Message)" -RawMessage $_.Exception.Message
                }
            } elseif ($RestoreItemAuthorEditor -and $skipFolderAuthorRestore) {
                Write-Log "DEBUG" "Restauration auteur dossier ignoree (chemin technique/long) : $($Item.RelPath)"
            }

            if ($CopyItemMetadata -and $dstFolderId) {
                Copy-DriveItemMetadata -SrcDriveId $SrcDriveId -SrcItemId $Item.Id -DstDriveId $DstDriveId -DstItemId $dstFolderId -SrcToken $SrcToken -DstToken $DstToken -ItemRelPath $Item.RelPath
            }

            if ($CopyPermissions) {
                try {
                    if ($dstFolderId) {
                        Copy-DriveItemPermissions -SrcDriveId $SrcDriveId -SrcItemId $Item.Id -DstDriveId $DstDriveId -DstItemId $dstFolderId -SrcToken $SrcToken -DstToken $DstToken -ItemRelPath $Item.RelPath
                    }
                } catch {
                    Write-Log "WARN" "Impossible de copier les permissions du dossier '$($Item.RelPath)' : $($_.Exception.Message)"
                }
            }
            return
        }

        if (-not $ForceOverwrite) {
            $segs     = ($Item.RelPath -replace '\\', '/') -split '/' | Where-Object { $_ } |
                        ForEach-Object { [Uri]::EscapeDataString($_) }
            $checkUri = "https://graph.microsoft.com/v1.0/drives/" + $DstDriveId + "/root:/" + ($segs -join '/')
            try {
                Invoke-GraphCall -Uri $checkUri -Token $DstToken | Out-Null
                Write-Log "INFO" "Ignore (deja present) : $($Item.RelPath)"
                return
            } catch { }
        }

        if ($script:RepairVersionHistoryOnlyResolved -and $RepairVersionPathRegex) {
            $p = ($Item.RelPath -replace '\\', '/')
            if ($p -notmatch $RepairVersionPathRegex) {
                Write-Log "DEBUG" "   ⏭  Repair-only : hors filtre regex, ignore ($($Item.RelPath))"
                return
            }
        }

        $parentPath = [System.IO.Path]::GetDirectoryName($Item.RelPath) -replace '\\', '/'
        if (-not $script:RepairVersionHistoryOnlyResolved -and $parentPath -and $parentPath -ne ".") {
            Ensure-DriveFolder -DriveId $DstDriveId -FolderRelPath $parentPath -Token $DstToken
        }

        if ($script:CopyModeResolved -eq "incremental" -and -not $CopyVersionHistory) {
            $segsInc     = ($Item.RelPath -replace '\\', '/') -split '/' | Where-Object { $_ } |
                           ForEach-Object { [Uri]::EscapeDataString($_) }
            $checkIncUri = "https://graph.microsoft.com/v1.0/drives/" + $DstDriveId + "/root:/" + ($segsInc -join '/')
            try {
                $incItem = Invoke-GraphCall -Uri $checkIncUri -Token $DstToken
                if ($incItem -and $incItem.id) {
                    Write-Log "INFO" "   ⏭  Incremental : deja present, ignore ($($Item.RelPath))"
                    return
                }
            } catch { }
        }

        # Resolve destination state before version replay
        $dstExistingId = $null
        $dstVersionCount = 0
        if ($CopyVersionHistory) {
            $segs     = ($Item.RelPath -replace '\\', '/') -split '/' | Where-Object { $_ } | ForEach-Object { [Uri]::EscapeDataString($_) }
            $checkUri = "https://graph.microsoft.com/v1.0/drives/" + $DstDriveId + "/root:/" + ($segs -join '/')
            try {
                $existingItem = Invoke-GraphCall -Uri $checkUri -Token $DstToken
                if ($existingItem -and $existingItem.id) { $dstExistingId = [string]$existingItem.id }
            } catch { }
        }

        Write-Log "INFO" "   ⬇  $($Item.RelPath)"
        $srcMeta = $null
        try {
            $srcMeta = Get-DriveItemMeta -DriveId $SrcDriveId -ItemId $Item.Id -Token $SrcToken
        } catch {
            Write-Log "WARN" "Metadonnees source indisponibles pour '$($Item.RelPath)' : $($_.Exception.Message)"
        }

        if ($CopyVersionHistory) {
            $versions = @()
            try {
                $versions = @(Get-DriveFileVersions -DriveId $SrcDriveId -ItemId $Item.Id -Token $SrcToken)
            } catch {
                Write-Log "WARN" "Historique indisponible pour '$($Item.RelPath)' : $($_.Exception.Message). Copie simple appliquee."
            }

            if ($versions.Count -gt 0) {
                # Complete = exactly src_count versions (ValidateUpdateListItem does not create extra versions)
                if ($dstExistingId) {
                    $dstVers = @()
                    $autoNeedsRepair = $false
                    try {
                        $dstVers = @(Get-DriveFileVersions -DriveId $DstDriveId -ItemId $dstExistingId -Token $DstToken)
                        $dstVersionCount = $dstVers.Count
                    } catch {}

                    if ($script:RepairVersionHistoryOnlyResolved) {
                        $needsRepair = Test-NeedsVersionHistoryRepair -SrcVersions $versions -DstVersions $dstVers
                        if (-not $needsRepair) {
                            Write-Log "INFO" "   ⏭  Repair-only : historique deja correct, ignore ($($Item.RelPath))"
                            return
                        }
                        Write-Log "WARN" "   🛠  Repair-only : historique incoherent detecte, rejeu ($($Item.RelPath))"
                    }

                    if (-not $script:RepairVersionHistoryOnlyResolved -and $script:CopyModeResolved -eq "incremental" -and $AutoRepairVersionHistoryInIncremental) {
                        $autoNeedsRepair = Test-NeedsVersionHistoryRepair -SrcVersions $versions -DstVersions $dstVers
                        if ($autoNeedsRepair) {
                            Write-Log "WARN" "   🛠  Incremental auto-repair : historique incoherent detecte, rejeu ($($Item.RelPath))"
                        }
                    }

                    if ($dstVersionCount -eq $versions.Count) {
                        if (-not $ReplayWhenVersionCountMatches -and -not $script:RepairVersionHistoryOnlyResolved -and -not $autoNeedsRepair) {
                            if ($CopyItemMetadata) {
                                Copy-DriveItemMetadata -SrcDriveId $SrcDriveId -SrcItemId $Item.Id -DstDriveId $DstDriveId -DstItemId $dstExistingId -SrcToken $SrcToken -DstToken $DstToken -ItemRelPath $Item.RelPath
                            }
                            if ($CopyPermissions) {
                                Copy-DriveItemPermissions -SrcDriveId $SrcDriveId -SrcItemId $Item.Id -DstDriveId $DstDriveId -DstItemId $dstExistingId -SrcToken $SrcToken -DstToken $DstToken -ItemRelPath $Item.RelPath
                            }
                            Write-Log "INFO" "   ⏭  Versions deja copiees ($dstVersionCount) : $($Item.RelPath)"
                            Write-Log "SUCCESS" "✅ $($Item.RelPath) - versions deja presentes"
                            return
                        }
                        if ($script:RepairVersionHistoryOnlyResolved) {
                            Write-Log "WARN" "   🔁 Repair-only : rejeu force (compte identique mais historique invalide) : $($Item.RelPath)"
                        } elseif ($autoNeedsRepair) {
                            Write-Log "WARN" "   🔁 Incremental auto-repair : rejeu force (compte identique mais historique invalide) : $($Item.RelPath)"
                        } else {
                            Write-Log "WARN" "   🔁 Rejeu force des versions (compte identique) : $($Item.RelPath)"
                        }
                    }
                    try {
                        $deleteUri = "https://graph.microsoft.com/v1.0/drives/" + $DstDriveId + "/items/" + $dstExistingId
                        Invoke-GraphCall -Uri $deleteUri -Token $DstToken -Method "DELETE" | Out-Null
                        if ($dstVersionCount -eq $versions.Count) {
                            Write-Log "DEBUG" "   🗑  Fichier destination supprime (rejeu force) : $($Item.RelPath)"
                        } else {
                            Write-Log "DEBUG" "   🗑  Fichier destination supprime (versions: $dstVersionCount != $($versions.Count)) : $($Item.RelPath)"
                        }
                    } catch { Write-Log "WARN" "Impossible de supprimer le fichier destination : $($_.Exception.Message)" }
                } elseif ($script:RepairVersionHistoryOnlyResolved) {
                    Write-Log "INFO" "   ⏭  Repair-only : fichier absent en destination, ignore ($($Item.RelPath))"
                    return
                }

                Write-Log "INFO" "   ⬆  Rejeu de $($versions.Count) version(s) : $($Item.RelPath)"
                $vCounter = 0
                $dstItemId = $null
                $dstListItemId = $null
                $dstListTitle = [System.IO.Path]::GetFileName($script:DstSiteInfo.LibraryPath)

                # Resolve createdDateTime once from source item metadata
                $srcCreated = $null
                if ($srcMeta) {
                    try {
                        if ($srcMeta.fileSystemInfo.createdDateTime) { $srcCreated = [datetime]$srcMeta.fileSystemInfo.createdDateTime }
                        elseif ($srcMeta.createdDateTime) { $srcCreated = [datetime]$srcMeta.createdDateTime }
                    } catch {}
                }
                $defaultVersionEmail = $null
                try {
                    if ($srcMeta -and $srcMeta.lastModifiedBy -and $srcMeta.lastModifiedBy.user -and $srcMeta.lastModifiedBy.user.email) {
                        $defaultVersionEmail = Normalize-Email -Value ([string]$srcMeta.lastModifiedBy.user.email)
                    }
                } catch {}
                $fallbackSourceEmail = Normalize-Email -Value $script:CurrentSourceUpn
                $fallbackTargetEmail = Normalize-Email -Value $script:CurrentTargetUpn

                foreach ($v in $versions) {
                    $vCounter++
                    $versionId = ""
                    $versionWhen = ""
                    $versionBy = ""
                    $versionActor = ""
                    try { if ($v.id) { $versionId = [string]$v.id } } catch {}
                    try { if ($v.lastModifiedDateTime) { $versionWhen = [string]$v.lastModifiedDateTime } } catch {}
                    try {
                        if ($v.lastModifiedBy.user.email) { $versionBy = [string]$v.lastModifiedBy.user.email }
                        if ($v.lastModifiedBy.user.displayName) { $versionActor = [string]$v.lastModifiedBy.user.displayName }
                        elseif ($v.lastModifiedBy.application.displayName) { $versionActor = [string]$v.lastModifiedBy.application.displayName }
                    } catch {}
                    if ([string]::IsNullOrWhiteSpace($versionId)) { continue }

                    if ($ExportVersionMetadata) {
                        Append-VersionMeta -FilePath $Item.RelPath -VersionId $versionId -ModifiedUtc $versionWhen -ModifiedBy $versionBy
                    }

                    # The current (last) version cannot be fetched via the versions endpoint
                    $vSize = Copy-DriveFileContent -SrcDriveId $SrcDriveId -SrcItemId $Item.Id -DstDriveId $DstDriveId -RelPath $Item.RelPath `
                        -SrcToken $SrcToken -DstToken $DstToken -VersionId $versionId -UseCurrentContentForVersion:($vCounter -eq $versions.Count) `
                        -ChunkBytes ($ChunkSizeMB * 1024 * 1024)
                    $actorForLog = if ($versionBy) { $versionBy } elseif ($versionActor) { $versionActor } else { "inconnu" }
                    Write-Log "DEBUG" "      Version $vCounter/$($versions.Count) rejouee (id=$versionId, par=$actorForLog, $vSize o)"

                    # Resolve dst item ID once after first upload
                    if (-not $dstItemId) {
                        $dstMeta = Get-DriveItemByRelPath -DriveId $DstDriveId -RelPath $Item.RelPath -Token $DstToken
                        if ($dstMeta -and $dstMeta.id) { $dstItemId = [string]$dstMeta.id }
                        if ($dstItemId) {
                            try { $dstListItemId = Get-DriveItemListItemId -DriveId $DstDriveId -ItemId $dstItemId -Token $DstToken } catch {}
                        }
                    }

                    if ($dstItemId) {
                        # Capture upload version ID before any PATCH/VVI changes it
                        $preOpVerId = $null
                        if ($PruneReplayDuplicateVersions) {
                            try {
                                $preOpVers = @(Get-DriveFileVersions -DriveId $DstDriveId -ItemId $dstItemId -Token $DstToken)
                                if ($preOpVers.Count -gt 0) { $preOpVerId = [string]($preOpVers | Select-Object -Last 1).id }
                            } catch {}
                        }

                        $vDate        = if ($versionWhen) { try { [datetime]$versionWhen } catch { $null } } else { $null }
                        $createdToUse = if ($srcCreated) { $srcCreated } elseif ($vDate) { $vDate } else { $null }
                        $srcVersionEmail = Normalize-Email -Value $versionBy
                        if (-not $srcVersionEmail -and $defaultVersionEmail) { $srcVersionEmail = $defaultVersionEmail }
                        if (-not $srcVersionEmail -and $fallbackSourceEmail) { $srcVersionEmail = $fallbackSourceEmail }
                        $mappedEmail = $null
                        if ($srcVersionEmail -and $script:UserMapping.ContainsKey($srcVersionEmail)) {
                            $mappedEmail = Normalize-Email -Value $script:UserMapping[$srcVersionEmail]
                        } elseif ($srcVersionEmail) {
                            $mappedEmail = $srcVersionEmail
                        }
                        if (-not $mappedEmail -and $fallbackTargetEmail) { $mappedEmail = $fallbackTargetEmail }

                        $editorDone = $false
                        $dateDone   = $false

                        # Primary for history replay: set Editor + Modified in one VVI call.
                        # This avoids creating an extra version with PATCH on fileSystemInfo.
                        $skipVersionAuthorRestore = Test-ShouldSkipAuthorRestore -ItemRelPath $Item.RelPath -ForVersion
                        if ($mappedEmail -and -not $skipVersionAuthorRestore -and $script:DstSpToken -and $script:DstSiteInfo) {
                            try {
                                $serverRelPath = Get-ServerRelativePathForItem -LibraryPath $script:DstSiteInfo.LibraryPath -ItemRelPath $Item.RelPath
                                $ensuredUserId = $null
                                try { $ensuredUserId = Invoke-SharePointEnsureUser -SpToken $script:DstSpToken -HostUrl $script:DstSiteInfo.HostUrl -SitePath $script:DstSiteInfo.SitePath -Email $mappedEmail } catch { Write-AuthorRestoreIssue -ItemRelPath $Item.RelPath -Text "      v$vCounter EnsureUser echoue ($mappedEmail) : $($_.Exception.Message)" -RawMessage $_.Exception.Message }
                                $peopleState = Set-SharePointAuthorEditor -SpToken $script:DstSpToken -HostUrl $script:DstSiteInfo.HostUrl `
                                    -SitePath $script:DstSiteInfo.SitePath -ServerRelPath $serverRelPath `
                                    -EditorEmail $mappedEmail -ModifiedDate $versionWhen -ListTitle $dstListTitle -ListItemId $dstListItemId

                                $verifiedEditor = $null
                                try { $verifiedEditor = Normalize-Email -Value ([string]$peopleState.Editor) } catch {}

                                if ($verifiedEditor -and $verifiedEditor -eq (Normalize-Email -Value $mappedEmail)) {
                                    $editorDone = $true
                                }

                                if ($versionWhen -and $peopleState -and $peopleState.Modified) {
                                    try {
                                        $srcWhen = ([datetime]$versionWhen).ToUniversalTime()
                                        $dstWhen = ([datetime]$peopleState.Modified).ToUniversalTime()
                                        $deltaSec = [math]::Abs(($dstWhen - $srcWhen).TotalSeconds)
                                        if ($deltaSec -le 120) { $dateDone = $true }
                                    } catch {}
                                }

                                if ($editorDone -and ($dateDone -or -not $versionWhen)) {
                                    Write-Log "DEBUG" "      v$vCounter : auteur+date via VVI OK ($versionBy -> $mappedEmail / $versionWhen)"
                                } else {
                                    Write-AuthorRestoreIssue -ItemRelPath $Item.RelPath -Text "      v$vCounter : auteur/date VVI non verifies ($versionBy -> $mappedEmail ; obtenu='$verifiedEditor' / modifie='$($peopleState.Modified)')"
                                    if ($EnableEditorIdFallback -and $ensuredUserId) {
                                        try {
                                            $mergeState = Set-SharePointAuthorEditorByLookupId -SpToken $script:DstSpToken -HostUrl $script:DstSiteInfo.HostUrl `
                                                -SitePath $script:DstSiteInfo.SitePath -ServerRelPath $serverRelPath -EditorLookupId $ensuredUserId `
                                                -ListTitle $dstListTitle -ListItemId $dstListItemId

                                            $verifiedEditorMerge = $null
                                            try { $verifiedEditorMerge = Normalize-Email -Value ([string]$mergeState.Editor) } catch {}
                                            if ($verifiedEditorMerge -and $verifiedEditorMerge -eq (Normalize-Email -Value $mappedEmail)) {
                                                Write-Log "DEBUG" "      v$vCounter : auteur via MERGE EditorId OK ($mappedEmail)"
                                                $editorDone = $true
                                            } else {
                                                Write-AuthorRestoreIssue -ItemRelPath $Item.RelPath -Text "      v$vCounter : auteur MERGE non verifie ($mappedEmail ; obtenu='$verifiedEditorMerge')"
                                            }
                                        } catch {
                                            Write-AuthorRestoreIssue -ItemRelPath $Item.RelPath -Text "      v$vCounter MERGE EditorId echoue : $($_.Exception.Message)" -RawMessage $_.Exception.Message
                                        }
                                    }
                                }
                            } catch {
                                Write-AuthorRestoreIssue -ItemRelPath $Item.RelPath -Text "      v$vCounter VVI echoue : $($_.Exception.Message)" -RawMessage $_.Exception.Message
                            }
                        } elseif ($mappedEmail -and $skipVersionAuthorRestore) {
                            Write-Log "DEBUG" "      v$vCounter : restauration auteur version ignoree (chemin technique/long)"
                        }

                        # No PATCH date in version-replay mode: PATCH creates duplicate versions.

                        # Delete the upload version if a new one was created by VVI or PATCH
                        if ($PruneReplayDuplicateVersions -and ($editorDone -or $dateDone) -and $preOpVerId) {
                            try {
                                $postVers  = @(Get-DriveFileVersions -DriveId $DstDriveId -ItemId $dstItemId -Token $DstToken)
                                $latestId  = if ($postVers.Count -gt 0) { [string]($postVers | Select-Object -Last 1).id } else { $null }
                                if ($latestId -and $latestId -ne $preOpVerId) {
                                    $delOpUri = "https://graph.microsoft.com/v1.0/drives/" + $DstDriveId + "/items/" + $dstItemId + "/versions/" + $preOpVerId
                                    Invoke-GraphCall -Uri $delOpUri -Token $DstToken -Method "DELETE" | Out-Null
                                    Write-Log "DEBUG" "      v$vCounter : doublon $preOpVerId supprime"
                                }
                            } catch {}
                        }

                        if (-not $mappedEmail -and $versionWhen) {
                            Write-AuthorRestoreIssue -ItemRelPath $Item.RelPath -Text "      v$vCounter : auteur non restaurable (email source introuvable). Acteur source='$actorForLog'"
                        } elseif (-not $editorDone) {
                            Write-AuthorRestoreIssue -ItemRelPath $Item.RelPath -Text "      v$vCounter : auteur non restaure (VVI non valide) pour '$mappedEmail'"
                        }
                    }
                }

                if ($script:RepairVersionHistoryOnlyResolved) {
                    Write-Log "SUCCESS" "✅ $($Item.RelPath) - historique repare"
                    return
                }

                if ($CopyItemMetadata) {
                    if (-not $dstItemId) {
                        try {
                            $dstMeta = Get-DriveItemByRelPath -DriveId $DstDriveId -RelPath $Item.RelPath -Token $DstToken
                            if ($dstMeta -and $dstMeta.id) { $dstItemId = [string]$dstMeta.id }
                        } catch {}
                    }
                    if ($dstItemId) {
                        Copy-DriveItemMetadata -SrcDriveId $SrcDriveId -SrcItemId $Item.Id -DstDriveId $DstDriveId -DstItemId $dstItemId -SrcToken $SrcToken -DstToken $DstToken -ItemRelPath $Item.RelPath
                    }
                }

                if ($CopyPermissions) {
                    if (-not $dstItemId) {
                        try {
                            $dstMeta = Get-DriveItemByRelPath -DriveId $DstDriveId -RelPath $Item.RelPath -Token $DstToken
                            if ($dstMeta -and $dstMeta.id) { $dstItemId = [string]$dstMeta.id }
                        } catch {}
                    }
                    if ($dstItemId) {
                        Copy-DriveItemPermissions -SrcDriveId $SrcDriveId -SrcItemId $Item.Id -DstDriveId $DstDriveId -DstItemId $dstItemId -SrcToken $SrcToken -DstToken $DstToken -ItemRelPath $Item.RelPath
                    }
                }

                $skipFileAuthorRestore = Test-ShouldSkipAuthorRestore -ItemRelPath $Item.RelPath
                if ($RestoreItemAuthorEditor -and -not $skipFileAuthorRestore -and $srcMeta -and $dstItemId -and $script:DstSpToken -and $script:DstSiteInfo) {
                    try {
                        $srcAuthor = $null
                        $srcEditor = $null
                        try { if ($srcMeta.createdBy.user.email) { $srcAuthor = [string]$srcMeta.createdBy.user.email } } catch {}
                        try { if ($srcMeta.lastModifiedBy.user.email) { $srcEditor = [string]$srcMeta.lastModifiedBy.user.email } } catch {}
                        if (-not $srcEditor) { $srcEditor = $srcAuthor }

                        $mappedAuthor = Resolve-MappedEmail -SourceEmail $srcAuthor -FallbackTargetEmail $script:CurrentTargetUpn
                        $mappedEditor = Resolve-MappedEmail -SourceEmail $srcEditor -FallbackTargetEmail $script:CurrentTargetUpn
                        $createdWhen  = if ($srcMeta.createdDateTime) { [string]$srcMeta.createdDateTime } else { "" }
                        $modifiedWhen = if ($srcMeta.lastModifiedDateTime) { [string]$srcMeta.lastModifiedDateTime } else { "" }

                        if ($mappedAuthor -or $mappedEditor) {
                            $authorId = $null
                            $editorId = $null
                            if ($mappedAuthor) { try { $authorId = Invoke-SharePointEnsureUser -SpToken $script:DstSpToken -HostUrl $script:DstSiteInfo.HostUrl -SitePath $script:DstSiteInfo.SitePath -Email $mappedAuthor } catch {} }
                            if ($mappedEditor) { try { $editorId = Invoke-SharePointEnsureUser -SpToken $script:DstSpToken -HostUrl $script:DstSiteInfo.HostUrl -SitePath $script:DstSiteInfo.SitePath -Email $mappedEditor } catch {} }

                            $serverRelPath = Get-ServerRelativePathForItem -LibraryPath $script:DstSiteInfo.LibraryPath -ItemRelPath $Item.RelPath
                            $peopleState = Set-SharePointAuthorEditor -SpToken $script:DstSpToken -HostUrl $script:DstSiteInfo.HostUrl -SitePath $script:DstSiteInfo.SitePath -ServerRelPath $serverRelPath -AuthorEmail $mappedAuthor -EditorEmail $mappedEditor -CreatedDate $createdWhen -ModifiedDate $modifiedWhen

                            $authorOk = $false
                            $editorOk = $false
                            try { $authorOk = ($peopleState.Author -eq (Normalize-Email -Value $mappedAuthor)) } catch {}
                            try { $editorOk = ($peopleState.Editor -eq (Normalize-Email -Value $mappedEditor)) } catch {}

                            if (-not $authorOk -or -not $editorOk) {
                                $mergeState = Set-SharePointAuthorEditorByLookupId -SpToken $script:DstSpToken -HostUrl $script:DstSiteInfo.HostUrl -SitePath $script:DstSiteInfo.SitePath -ServerRelPath $serverRelPath -AuthorLookupId $authorId -EditorLookupId $editorId
                                try { $authorOk = ($mergeState.Author -eq (Normalize-Email -Value $mappedAuthor)) } catch {}
                                try { $editorOk = ($mergeState.Editor -eq (Normalize-Email -Value $mappedEditor)) } catch {}
                            }

                            if ($authorOk -and $editorOk) {
                                Write-Log "DEBUG" "Auteur/Modifie par restaures (fichier) : $($Item.RelPath)"
                            } else {
                                Write-AuthorRestoreIssue -ItemRelPath $Item.RelPath -Text "Auteur/Modifie par non verifies (fichier) : $($Item.RelPath)"
                            }
                        }
                    } catch {
                        Write-AuthorRestoreIssue -ItemRelPath $Item.RelPath -Text "Restauration auteur/createur fichier echouee '$($Item.RelPath)' : $($_.Exception.Message)" -RawMessage $_.Exception.Message
                    }
                } elseif ($RestoreItemAuthorEditor -and $skipFileAuthorRestore) {
                    Write-Log "DEBUG" "Restauration auteur fichier ignoree (chemin technique/long) : $($Item.RelPath)"
                }

                Write-Log "INFO" "Restauration auteur version: tentative via SharePoint VVI terminee (voir logs vN pour le detail)."
                Write-Log "SUCCESS" "✅ $($Item.RelPath) - historique copie ($($versions.Count) version(s))"
                return
            }
        }

        Write-Log "INFO" "   ⬆  $($Item.RelPath)"
        $finalSize = Copy-DriveFileContent -SrcDriveId $SrcDriveId -SrcItemId $Item.Id -DstDriveId $DstDriveId -RelPath $Item.RelPath `
            -SrcToken $SrcToken -DstToken $DstToken -ChunkBytes ($ChunkSizeMB * 1024 * 1024)

        $dstMetaAfterUpload = $null
        try {
            $dstMetaAfterUpload = Get-DriveItemByRelPath -DriveId $DstDriveId -RelPath $Item.RelPath -Token $DstToken
        } catch {}

        if ($PreserveFileTimestamps -and $srcMeta) {
            try {
                $dstMeta = $dstMetaAfterUpload
                if ($dstMeta -and $dstMeta.id) {
                    $srcCreated = $null
                    $srcModified = $null
                    try {
                        if ($srcMeta.fileSystemInfo.createdDateTime) { $srcCreated = [datetime]$srcMeta.fileSystemInfo.createdDateTime }
                        elseif ($srcMeta.createdDateTime) { $srcCreated = [datetime]$srcMeta.createdDateTime }
                    } catch {}
                    try {
                        if ($srcMeta.fileSystemInfo.lastModifiedDateTime) { $srcModified = [datetime]$srcMeta.fileSystemInfo.lastModifiedDateTime }
                        elseif ($srcMeta.lastModifiedDateTime) { $srcModified = [datetime]$srcMeta.lastModifiedDateTime }
                    } catch {}

                    if ($srcCreated -and $srcModified) {
                        Set-DriveItemFileTimestamps -DriveId $DstDriveId -ItemId ([string]$dstMeta.id) -Token $DstToken -CreatedUtc $srcCreated -ModifiedUtc $srcModified
                        Write-Log "DEBUG" "Horodatages reappliques : $($Item.RelPath)"
                    }
                }
            } catch {
                Write-Log "WARN" "Impossible de reappliquer les horodatages pour '$($Item.RelPath)' : $($_.Exception.Message)"
            }
        }

        if ($CopyItemMetadata) {
            try {
                $dstMetaFields = $dstMetaAfterUpload
                if ($dstMetaFields -and $dstMetaFields.id) {
                    Copy-DriveItemMetadata -SrcDriveId $SrcDriveId -SrcItemId $Item.Id -DstDriveId $DstDriveId -DstItemId ([string]$dstMetaFields.id) -SrcToken $SrcToken -DstToken $DstToken -ItemRelPath $Item.RelPath
                }
            } catch {
                Write-Log "WARN" "Impossible de copier les metadonnees du fichier '$($Item.RelPath)' : $($_.Exception.Message)"
            }
        }

        if ($CopyVersionHistory) {
            Write-Log "INFO" "Restauration auteur version non applicable ici (copie simple sans rejeu d'historique)."
        }

        if ($CopyPermissions) {
            try {
                $dstMetaPerm = $dstMetaAfterUpload
                if ($dstMetaPerm -and $dstMetaPerm.id) {
                    Copy-DriveItemPermissions -SrcDriveId $SrcDriveId -SrcItemId $Item.Id -DstDriveId $DstDriveId -DstItemId ([string]$dstMetaPerm.id) -SrcToken $SrcToken -DstToken $DstToken -ItemRelPath $Item.RelPath
                }
            } catch {
                Write-Log "WARN" "Impossible de copier les permissions du fichier '$($Item.RelPath)' : $($_.Exception.Message)"
            }
        }

        $skipSimpleFileAuthorRestore = Test-ShouldSkipAuthorRestore -ItemRelPath $Item.RelPath
        if ($RestoreItemAuthorEditor -and -not $skipSimpleFileAuthorRestore -and $srcMeta -and $dstMetaAfterUpload -and $dstMetaAfterUpload.id -and $script:DstSpToken -and $script:DstSiteInfo) {
            try {
                $srcAuthor = $null
                $srcEditor = $null
                try { if ($srcMeta.createdBy.user.email) { $srcAuthor = [string]$srcMeta.createdBy.user.email } } catch {}
                try { if ($srcMeta.lastModifiedBy.user.email) { $srcEditor = [string]$srcMeta.lastModifiedBy.user.email } } catch {}
                if (-not $srcEditor) { $srcEditor = $srcAuthor }

                $mappedAuthor = Resolve-MappedEmail -SourceEmail $srcAuthor -FallbackTargetEmail $script:CurrentTargetUpn
                $mappedEditor = Resolve-MappedEmail -SourceEmail $srcEditor -FallbackTargetEmail $script:CurrentTargetUpn
                $createdWhen  = if ($srcMeta.createdDateTime) { [string]$srcMeta.createdDateTime } else { "" }
                $modifiedWhen = if ($srcMeta.lastModifiedDateTime) { [string]$srcMeta.lastModifiedDateTime } else { "" }

                if ($mappedAuthor -or $mappedEditor) {
                    $authorId = $null
                    $editorId = $null
                    if ($mappedAuthor) { try { $authorId = Invoke-SharePointEnsureUser -SpToken $script:DstSpToken -HostUrl $script:DstSiteInfo.HostUrl -SitePath $script:DstSiteInfo.SitePath -Email $mappedAuthor } catch {} }
                    if ($mappedEditor) { try { $editorId = Invoke-SharePointEnsureUser -SpToken $script:DstSpToken -HostUrl $script:DstSiteInfo.HostUrl -SitePath $script:DstSiteInfo.SitePath -Email $mappedEditor } catch {} }

                    $serverRelPath = Get-ServerRelativePathForItem -LibraryPath $script:DstSiteInfo.LibraryPath -ItemRelPath $Item.RelPath
                    $peopleState = Set-SharePointAuthorEditor -SpToken $script:DstSpToken -HostUrl $script:DstSiteInfo.HostUrl -SitePath $script:DstSiteInfo.SitePath -ServerRelPath $serverRelPath -AuthorEmail $mappedAuthor -EditorEmail $mappedEditor -CreatedDate $createdWhen -ModifiedDate $modifiedWhen

                    $authorOk = $false
                    $editorOk = $false
                    try { $authorOk = ($peopleState.Author -eq (Normalize-Email -Value $mappedAuthor)) } catch {}
                    try { $editorOk = ($peopleState.Editor -eq (Normalize-Email -Value $mappedEditor)) } catch {}

                    if (-not $authorOk -or -not $editorOk) {
                        $mergeState = Set-SharePointAuthorEditorByLookupId -SpToken $script:DstSpToken -HostUrl $script:DstSiteInfo.HostUrl -SitePath $script:DstSiteInfo.SitePath -ServerRelPath $serverRelPath -AuthorLookupId $authorId -EditorLookupId $editorId
                        try { $authorOk = ($mergeState.Author -eq (Normalize-Email -Value $mappedAuthor)) } catch {}
                        try { $editorOk = ($mergeState.Editor -eq (Normalize-Email -Value $mappedEditor)) } catch {}
                    }

                    if ($authorOk -and $editorOk) {
                        Write-Log "DEBUG" "Auteur/Modifie par restaures (fichier simple) : $($Item.RelPath)"
                    } else {
                        Write-AuthorRestoreIssue -ItemRelPath $Item.RelPath -Text "Auteur/Modifie par non verifies (fichier simple) : $($Item.RelPath)"
                    }
                }
            } catch {
                Write-AuthorRestoreIssue -ItemRelPath $Item.RelPath -Text "Restauration auteur/createur fichier simple echouee '$($Item.RelPath)' : $($_.Exception.Message)" -RawMessage $_.Exception.Message
            }
        } elseif ($RestoreItemAuthorEditor -and $skipSimpleFileAuthorRestore) {
            Write-Log "DEBUG" "Restauration auteur fichier simple ignoree (chemin technique/long) : $($Item.RelPath)"
        }

        Write-Log "SUCCESS" "✅ $($Item.RelPath) ($finalSize o)"

    } catch {
        Append-Err -FilePath $Item.RelPath -Msg $_.Exception.Message -Phase "copy"
        Write-Log "ERROR" "❌ $($Item.RelPath) : $($_.Exception.Message)"
    }
}

# =========================
# FONCTION DE MIGRATION D'UN UTILISATEUR
# =========================
function Invoke-MigrationUser {
    param(
        [string]$OneDriveUrlSource,
        [string]$OneDriveUrlCible,
        [string]$UpnSource,
        [string]$UpnCible,
        [int]$NumeroLigne
    )

    Write-Log "INFO"    "--------------------------------------------"
    Write-Log "SUCCESS" "=== MIGRATION LIGNE $NumeroLigne ==="
    Write-Log "INFO"    "Source UPN : $UpnSource"
    Write-Log "INFO"    "Cible  UPN : $UpnCible"
    Write-Log "INFO"    "--------------------------------------------"

    $mode = if ($CopyMode) { ([string]$CopyMode).Trim().ToLowerInvariant() } else { "full" }
    if ($mode -ne "full" -and $mode -ne "incremental") {
        Write-Log "WARN" "CopyMode invalide '$CopyMode'. Valeurs supportees: full|incremental. Mode 'full' applique."
        $mode = "full"
    }
    $script:CopyModeResolved = $mode
    Write-Log "INFO" "Mode de copie : $mode"
    if ($mode -eq "incremental") {
        Write-Log "INFO" "Mode incremental: seuls les fichiers absents en destination sont copies."
        if ($AutoRepairVersionHistoryInIncremental -and $CopyVersionHistory) {
            Write-Log "INFO" "Mode incremental auto-repair: historique versions detecte/corrige automatiquement si incoherent."
        }
    }

    $script:RepairVersionHistoryOnlyResolved = [bool]$RepairVersionHistoryOnly
    if ($script:RepairVersionHistoryOnlyResolved) {
        if (-not $CopyVersionHistory) {
            Write-Log "WARN" "RepairVersionHistoryOnly=true force CopyVersionHistory=true"
            $CopyVersionHistory = $true
        }
        Write-Log "INFO" "Mode repair-only: rejeu uniquement des fichiers dont l'historique destination est incoherent."
        if ($RepairVersionPathRegex) {
            Write-Log "INFO" "Filtre repair-only actif : $RepairVersionPathRegex"
        }
    }

    $script:CurrentSourceUpn = $UpnSource
    $script:CurrentTargetUpn = $UpnCible

    # Baseline mapping to restore owner versions even if not listed in user_mapping.csv.
    $srcUpnMap = Normalize-Email -Value $UpnSource
    $dstUpnMap = Normalize-Email -Value $UpnCible
    if ($srcUpnMap -and $dstUpnMap -and -not $script:UserMapping.ContainsKey($srcUpnMap)) {
        $script:UserMapping[$srcUpnMap] = $dstUpnMap
        Write-Log "DEBUG" "Mapping auto UPN ajoute : $srcUpnMap -> $dstUpnMap"
    }

    # Extraire le folder depuis les URLs
    $srcFolder = (([Uri]$OneDriveUrlSource).AbsolutePath -split '/')[2]
    $dstFolder = (([Uri]$OneDriveUrlCible).AbsolutePath  -split '/')[2]

    # Tokens
    $srcToken = Get-GraphToken -TenantId $app_source.tenant_id `
                -ClientId $app_source.client_id -ClientSecret $app_source.client_secret
    $dstToken = Get-AppToken -TenantId $app_cible.tenant_id -ClientId $app_cible.client_id.Trim() `
                -Scope "https://graph.microsoft.com/.default" `
                -ClientSecret $app_cible.client_secret `
                -CertThumbprint $app_cible.certificate_thumbprint `
                -CertPath $app_cible.certificate_path `
                -CertPassword $app_cible.certificate_password `
                -Label "CIBLE Graph"

    # Résolution des Drive IDs
    Write-Log "INFO" "=== RESOLUTION DES DRIVES ==="
    $srcDriveId = Get-DriveId -UPN $UpnSource -UserFolder $srcFolder -Token $srcToken -Label "SOURCE"
    $dstDriveId = Get-DriveId -UPN $UpnCible  -UserFolder $dstFolder -Token $dstToken -Label "CIBLE"
    Write-Log "SUCCESS" "Drive source : $srcDriveId"
    Write-Log "SUCCESS" "Drive cible  : $dstDriveId"

    if ($CopyPermissions) {
        try {
            $srcRoot = Get-DriveRootItem -DriveId $srcDriveId -Token $srcToken
            $dstRoot = Get-DriveRootItem -DriveId $dstDriveId -Token $dstToken
            if ($srcRoot -and $dstRoot -and $srcRoot.id -and $dstRoot.id) {
                Copy-DriveItemPermissions -SrcDriveId $srcDriveId -SrcItemId ([string]$srcRoot.id) -DstDriveId $dstDriveId -DstItemId ([string]$dstRoot.id) -SrcToken $srcToken -DstToken $dstToken -ItemRelPath "/"
                Write-Log "INFO" "Permissions racine OneDrive traitees"
            }
        } catch {
            Write-Log "WARN" "Impossible de copier les permissions racine : $($_.Exception.Message)"
        }
    }

    if ($CopyItemMetadata) {
        try {
            $srcRootMeta = Get-DriveRootItem -DriveId $srcDriveId -Token $srcToken
            $dstRootMeta = Get-DriveRootItem -DriveId $dstDriveId -Token $dstToken
            if ($srcRootMeta -and $dstRootMeta -and $srcRootMeta.id -and $dstRootMeta.id) {
                Copy-DriveItemMetadata -SrcDriveId $srcDriveId -SrcItemId ([string]$srcRootMeta.id) -DstDriveId $dstDriveId -DstItemId ([string]$dstRootMeta.id) -SrcToken $srcToken -DstToken $dstToken -ItemRelPath "/"
                Write-Log "INFO" "Metadonnees racine OneDrive ignorees (limitation Graph)"
            }
        } catch {
            Write-Log "WARN" "Impossible de copier les metadonnees racine : $($_.Exception.Message)"
        }
    }

    # Obtenir le token SharePoint REST et les infos du site destination pour la restauration des auteurs
    $script:DstSpToken  = $null
    $script:DstSiteInfo = $null
    try {
        $spHost = $null
        try {
            $drvInfo = Invoke-GraphCall -Uri ("https://graph.microsoft.com/v1.0/drives/" + $dstDriveId + "?`$select=webUrl") -Token $dstToken
            $spHost  = ([Uri]$drvInfo.webUrl).Scheme + "://" + ([Uri]$drvInfo.webUrl).Host
        } catch {}
        if ($spHost) {
            $script:DstSpToken  = Get-AppToken -TenantId $app_cible.tenant_id -ClientId $app_cible.client_id.Trim() `
                -Scope ($spHost.TrimEnd('/') + "/.default") `
                -ClientSecret $app_cible.client_secret `
                -CertThumbprint $app_cible.certificate_thumbprint `
                -CertPath $app_cible.certificate_path `
                -CertPassword $app_cible.certificate_password `
                -Label "CIBLE SharePoint"
            $script:DstSiteInfo = Get-DriveSiteInfo -DriveId $dstDriveId -Token $dstToken
            Write-Log "SUCCESS" "Token SharePoint REST obtenu pour $spHost"

            # Diagnostics: decode SP token roles and test a simple GET
            try {
                $parts = $script:DstSpToken -split '\.'
                if ($parts.Count -ge 2) {
                    $pad = $parts[1] -replace '-','+' -replace '_','/'
                    while ($pad.Length % 4) { $pad += '=' }
                    $payload = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($pad)) | ConvertFrom-Json
                    $roles = if ($payload.roles) { ($payload.roles) -join ', ' } else { '(aucun)' }
                    Write-Log "DEBUG" "SP token roles : $roles | aud : $($payload.aud)"
                }
            } catch {}
            try {
                $testUri = $spHost + $script:DstSiteInfo.SitePath + "/_api/web?`$select=Title"
                $testReq = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $testUri)
                $testReq.Headers.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $script:DstSpToken)
                $testReq.Headers.TryAddWithoutValidation("Accept", "application/json;odata=nometadata") | Out-Null
                $testResp = $script:HttpClient.SendAsync($testReq).Result
                $testBody = $testResp.Content.ReadAsStringAsync().Result
                Write-Log "DEBUG" "SP REST test GET /_api/web : HTTP $([int]$testResp.StatusCode) - $($testBody.Substring(0, [Math]::Min(200, $testBody.Length)))"
            } catch { Write-Log "DEBUG" "SP REST test echoue : $($_.Exception.Message)" }
        }
    } catch {
        Write-Log "WARN" "Token SharePoint REST non obtenu (auteur des versions non restaurable) : $($_.Exception.Message)"
    }

    # Scan
    Write-Log "INFO" "=== SCAN DU ONEDRIVE SOURCE ==="
    $srcToken = Get-GraphToken -TenantId $app_source.tenant_id `
                -ClientId $app_source.client_id -ClientSecret $app_source.client_secret
    $allItems  = Get-AllDriveItems -DriveId $srcDriveId -ItemId "root" -CurrentPath "" -Token $srcToken

    $folders   = @($allItems | Where-Object {  $_.IsFolder })
    $files     = @($allItems | Where-Object { -not $_.IsFolder })
    $totalSize = ($files | Measure-Object -Property Size -Sum).Sum

    Write-Log "SUCCESS" "Trouve : $($folders.Count) dossier(s) / $($files.Count) fichier(s)"
    Write-Log "INFO"    "Taille totale : $([math]::Round($totalSize / 1MB, 2)) MB"

    $resumeKey = Get-ResumeKey -UpnSource $UpnSource -UpnCible $UpnCible
    $resume = Get-ResumeProgress -ResumeKey $resumeKey
    $startFolderIndex = 0
    $startFileIndex = 0
    if ($resume) {
        $startFolderIndex = [Math]::Max(0, [int]$resume.FolderIndex)
        $startFileIndex   = [Math]::Max(0, [int]$resume.FileIndex)
        Write-Log "WARN" "Reprise active : dossiers=$startFolderIndex, fichiers=$startFileIndex"
    }

    # Dossiers
    Write-Log "INFO" "=== CREATION DES DOSSIERS ==="
    $sortedFolders = @($folders | Sort-Object { ($_.RelPath -split '/').Count }, RelPath)
    $srcTokenFolders = Get-GraphToken -TenantId $app_source.tenant_id `
                      -ClientId $app_source.client_id -ClientSecret $app_source.client_secret
    $dstTokenFolders = Get-AppToken -TenantId $app_cible.tenant_id -ClientId $app_cible.client_id.Trim() `
                      -Scope "https://graph.microsoft.com/.default" `
                      -ClientSecret $app_cible.client_secret `
                      -CertThumbprint $app_cible.certificate_thumbprint `
                      -CertPath $app_cible.certificate_path `
                      -CertPassword $app_cible.certificate_password

    for ($i = $startFolderIndex; $i -lt $sortedFolders.Count; $i++) {
        $folder = $sortedFolders[$i]
        $srcTokenFolders = Resolve-ActiveToken -Token $srcTokenFolders
        $dstTokenFolders = Resolve-ActiveToken -Token $dstTokenFolders
        $errBefore = $script:ErrRows.Count
        Copy-DriveItem -Item $folder -SrcDriveId $srcDriveId -DstDriveId $dstDriveId `
                       -SrcToken $srcTokenFolders -DstToken $dstTokenFolders
        Flush-PeriodicCsvSnapshots -ProcessedItems ($i + 1)
        if ($script:ErrRows.Count -eq $errBefore) {
            Set-ResumeProgress -ResumeKey $resumeKey -FolderIndex ($i + 1) -FileIndex $startFileIndex
            if ($EnableResumeCheckpoint -and ((($i + 1) % $CheckpointEveryItems) -eq 0)) {
                Save-ResumeState
            }
        }
    }

    # Fichiers
    Write-Log "INFO" "=== COPIE DES FICHIERS ==="
    $srcTokenFiles = Get-GraphToken -TenantId $app_source.tenant_id `
                    -ClientId $app_source.client_id -ClientSecret $app_source.client_secret
    $dstTokenFiles = Get-AppToken -TenantId $app_cible.tenant_id -ClientId $app_cible.client_id.Trim() `
                    -Scope "https://graph.microsoft.com/.default" `
                    -ClientSecret $app_cible.client_secret `
                    -CertThumbprint $app_cible.certificate_thumbprint `
                    -CertPath $app_cible.certificate_path `
                    -CertPassword $app_cible.certificate_password

    for ($i = $startFileIndex; $i -lt $files.Count; $i++) {
        $file = $files[$i]
        $counter = $i + 1
        Write-Log "INFO" "[$counter/$($files.Count)] $($file.RelPath) ($($file.Size) o)"
        $srcTokenFiles = Resolve-ActiveToken -Token $srcTokenFiles
        $dstTokenFiles = Resolve-ActiveToken -Token $dstTokenFiles
        $errBefore = $script:ErrRows.Count
        Copy-DriveItem -Item $file -SrcDriveId $srcDriveId -DstDriveId $dstDriveId `
                       -SrcToken $srcTokenFiles -DstToken $dstTokenFiles
        Flush-PeriodicCsvSnapshots -ProcessedItems ($i + 1)
        if ($script:ErrRows.Count -eq $errBefore) {
            Set-ResumeProgress -ResumeKey $resumeKey -FolderIndex $sortedFolders.Count -FileIndex ($i + 1)
            if ($EnableResumeCheckpoint -and ((($i + 1) % $CheckpointEveryItems) -eq 0)) {
                Save-ResumeState
            }
        }
    }

    Clear-ResumeProgress -ResumeKey $resumeKey
    Save-ResumeState

    $errMigration = $script:ErrRows.Count
    Write-Log "SUCCESS" "=== MIGRATION LIGNE $NumeroLigne TERMINEE ==="
    Write-Log "INFO"    "Total   : $($files.Count) fichier(s)"
    Write-Log "INFO"    "Erreurs : $errMigration"
}

# =========================
# SCRIPT PRINCIPAL
# =========================
try {
    Write-Log "INFO"    "============================================"
    Write-Log "SUCCESS" "=== COPIE ONEDRIVE CROSS-TENANT (GRAPH) ==="
    Write-Log "INFO"    "============================================"

    Load-ResumeState

    # Vérifier l'existence du CSV
    if (-not (Test-Path -LiteralPath $CsvPath)) {
        throw "Fichier CSV introuvable : $CsvPath"
    }

    # Charger le mapping utilisateurs si le fichier existe
    if (Test-Path -LiteralPath $UserMappingCsvPath) {
        try {
            $mappingRows = Import-Csv -LiteralPath $UserMappingCsvPath -Delimiter ';' -Encoding UTF8
            foreach ($row in $mappingRows) {
                $src = $row.PSObject.Properties.Value[0].Trim().ToLower()
                $dst = $row.PSObject.Properties.Value[1].Trim()
                if ($src -and $dst) { $script:UserMapping[$src] = $dst }
            }
            Write-Log "INFO" "Mapping utilisateurs charge : $($script:UserMapping.Count) entree(s) depuis $UserMappingCsvPath"
        } catch {
            Write-Log "WARN" "Impossible de charger le mapping utilisateurs : $($_.Exception.Message)"
        }
    } else {
        Write-Log "WARN" "Fichier de mapping utilisateurs absent ($UserMappingCsvPath) - auteurs non restaures"
    }

    # Lire le CSV (séparateur ;)
    # Colonnes attendues : onedrive_url_source ; onedrive_url_cible ; upn_source ; upn_cible
    $migrations = @(Import-Csv -LiteralPath $CsvPath -Delimiter ';' -Encoding UTF8)

    if ($migrations.Count -eq 0) {
        throw "Le CSV est vide : $CsvPath"
    }

    Write-Log "INFO" "CSV charge : $($migrations.Count) ligne(s) trouvee(s) dans $CsvPath"

    # Valider les colonnes
    $colonnesRequises = @('onedrive_url_source','onedrive_url_cible','upn_source','upn_cible')
    foreach ($col in $colonnesRequises) {
        if ($migrations[0].PSObject.Properties.Name -notcontains $col) {
            throw "Colonne manquante dans le CSV : '$col'. Colonnes attendues : $($colonnesRequises -join ', ')"
        }
    }

    # Boucle sur chaque ligne du CSV
    $ligneNum   = 0
    $okCount    = 0
    $errCount   = 0

    foreach ($row in $migrations) {
        $ligneNum++

        # Ignorer les lignes vides
        if ([string]::IsNullOrWhiteSpace($row.upn_source) -or [string]::IsNullOrWhiteSpace($row.upn_cible)) {
            Write-Log "WARN" "Ligne $ligneNum ignoree (UPN vide)"
            continue
        }

        try {
            Invoke-MigrationUser `
                -OneDriveUrlSource $row.onedrive_url_source.Trim() `
                -OneDriveUrlCible  $row.onedrive_url_cible.Trim() `
                -UpnSource         $row.upn_source.Trim() `
                -UpnCible          $row.upn_cible.Trim() `
                -NumeroLigne       $ligneNum

            $okCount++
        } catch {
            $errCount++
            Write-Log "ERROR" "Ligne $ligneNum ECHOUEE : $($_.Exception.Message)"
            Append-Err -FilePath "migration_ligne_$ligneNum" -Msg $_.Exception.Message -Phase "migration"
        }
    }

    Write-Log "SUCCESS" "============================================"
    Write-Log "SUCCESS" "=== TOUTES LES MIGRATIONS TERMINEES ==="
    Write-Log "INFO"    "Total lignes  : $ligneNum"
    Write-Log "INFO"    "Succes        : $okCount"
    Write-Log "INFO"    "Echecs        : $errCount"
    Write-Log "SUCCESS" "============================================"

} catch {
    Write-Log "ERROR" "ERREUR FATALE : $($_.Exception.Message)"
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    exit 1
} finally {
    if ($ExportVersionMetadata -and $script:VersionRows.Count -gt 0) {
        try {
            $script:VersionRows | Export-Csv -LiteralPath $script:VersionCsvPath -Encoding UTF8 -NoTypeInformation
            Write-Log "INFO" "Metadonnees versions exportees : $script:VersionCsvPath"
        } catch { }
    }
    if ($script:ErrRows.Count -gt 0) {
        try {
            $script:ErrRows | Export-Csv -LiteralPath $script:ErrCsvPath -Encoding UTF8 -NoTypeInformation
            Write-Log "WARN" "$($script:ErrRows.Count) erreur(s) : $script:ErrCsvPath"
        } catch { }
    }
    if ($script:HttpClient) { $script:HttpClient.Dispose() }
    Write-Log "INFO" "Log : $script:LogPath"
}