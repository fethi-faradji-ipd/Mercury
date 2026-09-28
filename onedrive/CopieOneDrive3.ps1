# =========================
# CONFIGURATION
# =========================

# App Registration TENANT SOURCE
$app_source = @{
tenant_Id = "85a5a352-25d6-4894-a97d-221cd1712dd2"
        client_Id = "684f670e-6931-4120-b053-f8c9538744f6"
        client_Secret = " "
}

# App Registration TENANT CIBLE
$app_cible = @{
    tenant_id     = "2eea08b8-1972-447b-ad43-d044d042500a"
    client_id     = " 821109a4-6e0e-48e8-b477-8f9b70aa32a4"
    client_secret = " "  # laisser vide si utilisation du certificat
    # Authentification par certificat (recommande pour SharePoint REST)
    # Renseigner certificate_thumbprint (certificat dans le store Windows Cert:\CurrentUser\My)
    # OU certificate_path (chemin vers fichier .pfx)
    certificate_thumbprint = ""  # ex : "A1B2C3D4E5F6..."
    certificate_path       = ""  # ex : "C:\certs\app_cible.pfx"
    certificate_password   = ""  # mot de passe du PFX si besoin
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
$ForceOverwrite = $true
$ChunkSizeMB    = 10
$CopyVersionHistory = $true
$MaxVersionsPerFile = 0
$PreserveFileTimestamps = $true
$ExportVersionMetadata  = $true
$CopyPermissions        = $true
$CopyItemMetadata       = $true

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
$script:DstSpToken  = $null
$script:DstSiteInfo = $null
$script:UserMapping = @{}  # source_email -> destination_email

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
        $pwd   = if ($CertPassword) { $CertPassword } else { [string]::Empty }
        $cert  = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertPath, $pwd, $flags)
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
        [string]$Label              = ""
    )
    $cacheKey = "$TenantId|$ClientId|$Scope"
    $now      = Get-Date
    if ($script:TokenCache.ContainsKey($cacheKey)) {
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
        [Parameter(Mandatory)][string]$ClientSecret
    )
    $cacheKey = "$TenantId|$ClientId|GRAPH"
    $now      = Get-Date
    $skew     = [TimeSpan]::FromMinutes(5)

    if ($script:TokenCache.ContainsKey($cacheKey)) {
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
        [Parameter(Mandatory)][string]$SharePointHost
    )
    $scope    = $SharePointHost.TrimEnd('/') + "/.default"
    $cacheKey = "$TenantId|$ClientId|SP|$SharePointHost"
    $now      = Get-Date
    if ($script:TokenCache.ContainsKey($cacheKey)) {
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
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        $req = [System.Net.Http.HttpRequestMessage]::new(
            [System.Net.Http.HttpMethod]::new($Method), $Uri)
        $req.Headers.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $Token)

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
        $formValues += @{
            FieldName  = "Modified"
            FieldValue = $ModifiedDate
        }
    }
    $body = @{
        formValues         = $formValues
        bNewDocumentUpdate = $true
        checkInComment     = ""
    } | ConvertTo-Json -Depth 10 -Compress
    $safePath = $ServerRelPath -replace "'", "''"
    $apiUrl = $HostUrl + $SitePath + "/_api/web/GetFileByServerRelativePath(decodedurl='" + $safePath + "')/ListItemAllFields/ValidateUpdateListItem()"
    $req    = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Post, $apiUrl)
    $req.Headers.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $SpToken)
    $req.Headers.TryAddWithoutValidation("Accept", "application/json;odata=nometadata") | Out-Null
    $req.Content = [System.Net.Http.StringContent]::new($body, [System.Text.Encoding]::UTF8, "application/json")
    $resp   = $script:HttpClient.SendAsync($req).Result
    $respBody = $resp.Content.ReadAsStringAsync().Result
    if (-not $resp.IsSuccessStatusCode) {
        throw "SP REST $([int]$resp.StatusCode) : $respBody"
    }
    # Log VVI response to show what fields were actually updated
    try {
        $vviResult = $respBody | ConvertFrom-Json
        $editorErr = ($vviResult.value | Where-Object { $_.FieldName -eq 'Editor' }).ErrorMessage
        $modErr    = ($vviResult.value | Where-Object { $_.FieldName -eq 'Modified' }).ErrorMessage
        Write-Log "DEBUG" "VVI reponse : Editor=$editorErr | Modified=$modErr"
    } catch {}
    # Read back the actual Editor after VVI to verify
    try {
        $getUri  = $HostUrl + $SitePath + "/_api/web/GetFileByServerRelativePath(decodedurl='" + ($ServerRelPath -replace "'","''") + "')/ListItemAllFields?`$select=Editor/Title,Editor/EMail,Modified&`$expand=Editor"
        $getReq  = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $getUri)
        $getReq.Headers.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $SpToken)
        $getReq.Headers.TryAddWithoutValidation("Accept", "application/json;odata=nometadata") | Out-Null
        $getResp = $script:HttpClient.SendAsync($getReq).Result
        $getBody = $getResp.Content.ReadAsStringAsync().Result | ConvertFrom-Json
        Write-Log "DEBUG" "VVI verification : Editor=$($getBody.Editor.EMail) / Modified=$($getBody.Modified)"
    } catch { Write-Log "DEBUG" "VVI verification echouee : $($_.Exception.Message)" }
}

# Provision user in the site and return their site user Id (EnsureUser — requires Sites.Manage.All on SharePoint)
function Invoke-SharePointEnsureUser {
    param([string]$SpToken, [string]$HostUrl, [string]$SitePath, [string]$Email)
    $body   = '{"logonName":"i:0#.f|membership|' + $Email + '"}'
    $apiUrl = $HostUrl + $SitePath + "/_api/web/EnsureUser"
    $req    = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Post, $apiUrl)
    $req.Headers.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $SpToken)
    $req.Headers.TryAddWithoutValidation("Accept", "application/json;odata=nometadata") | Out-Null
    $req.Content = [System.Net.Http.StringContent]::new($body, [System.Text.Encoding]::UTF8, "application/json")
    $resp   = $script:HttpClient.SendAsync($req).Result
    if (-not $resp.IsSuccessStatusCode) {
        $errBody = $resp.Content.ReadAsStringAsync().Result
        throw "EnsureUser $([int]$resp.StatusCode) : $errBody"
    }
    $json = $resp.Content.ReadAsStringAsync().Result | ConvertFrom-Json
    return [string]$json.Id
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

        Set-DriveItemListFields -DriveId $DstDriveId -ItemId $DstItemId -Fields $fieldsToCopy -Token $DstToken
        Write-Log "DEBUG" "Metadonnees appliquees : $ItemRelPath ($($fieldsToCopy.Count) champ(s))"
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

    $allPerms = New-Object System.Collections.Generic.List[object]
    $url = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/items/" + $ItemId + "/permissions?" + '$top=200'

    do {
        $resp = Invoke-GraphCall -Uri $url -Token $Token
        if ($resp -and $resp.value) {
            foreach ($p in $resp.value) { $allPerms.Add($p) | Out-Null }
        }
        $url = $resp.'@odata.nextLink'
    } while ($url)

    return @($allPerms)
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
        $encSegs     = ($currentPath -split '/') | ForEach-Object { [Uri]::EscapeDataString($_) }
        $encCurrent  = $encSegs -join '/'
        $checkUri    = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/root:/" + $encCurrent
        $exists      = $true

        try {
            Invoke-GraphCall -Uri $checkUri -Token $Token | Out-Null
        } catch { $exists = $false }

        if (-not $exists) {
            $parentSegs = ($currentPath -split '/')[0..($segs.Count - 2)]
            if ($parentSegs.Count -gt 0 -and $parentSegs[0]) {
                $encParent = ($parentSegs | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/'
                $createUri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/root:/" + $encParent + ":/children"
            } else {
                $createUri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/root/children"
            }
            $bodyJson = '{"name":"' + ($seg -replace '"', '\"') + '","folder":{},"@microsoft.graph.conflictBehavior":"replace"}'
            try {
                Invoke-GraphCall -Uri $createUri -Token $Token -Method "POST" -StringBody $bodyJson | Out-Null
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
    $segs      = ($RelPath -replace '\\', '/') -split '/' | Where-Object { $_ } |
                 ForEach-Object { [Uri]::EscapeDataString($_) }
    $encPath   = $segs -join '/'
    # Use session instead of PUT so eTag is not checked on subsequent version uploads
    $createUri = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/root:/" + $encPath + ":/createUploadSession"
    $session   = Invoke-GraphCall -Uri $createUri -Token $Token -Method "POST" `
                     -StringBody '{"item":{"@microsoft.graph.conflictBehavior":"replace"}}'
    $uploadUrl = $session.uploadUrl
    $total     = $Bytes.Length
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
            Ensure-DriveFolder -DriveId $DstDriveId -FolderRelPath $Item.RelPath -Token $DstToken
            $dstFolderId = $null
            try {
                $dstFolder = Get-DriveItemByRelPath -DriveId $DstDriveId -RelPath $Item.RelPath -Token $DstToken
                if ($dstFolder -and $dstFolder.id) { $dstFolderId = [string]$dstFolder.id }
            } catch {}

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

        $parentPath = [System.IO.Path]::GetDirectoryName($Item.RelPath) -replace '\\', '/'
        if ($parentPath -and $parentPath -ne ".") {
            Ensure-DriveFolder -DriveId $DstDriveId -FolderRelPath $parentPath -Token $DstToken
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
                    try {
                        $dstVers = @(Get-DriveFileVersions -DriveId $DstDriveId -ItemId $dstExistingId -Token $DstToken)
                        $dstVersionCount = $dstVers.Count
                    } catch {}
                    if ($dstVersionCount -eq $versions.Count) {
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
                    try {
                        $deleteUri = "https://graph.microsoft.com/v1.0/drives/" + $DstDriveId + "/items/" + $dstExistingId
                        Invoke-GraphCall -Uri $deleteUri -Token $DstToken -Method "DELETE" | Out-Null
                        Write-Log "DEBUG" "   🗑  Fichier destination supprime (versions: $dstVersionCount != $($versions.Count)) : $($Item.RelPath)"
                    } catch { Write-Log "WARN" "Impossible de supprimer le fichier destination : $($_.Exception.Message)" }
                }

                Write-Log "INFO" "   ⬆  Rejeu de $($versions.Count) version(s) : $($Item.RelPath)"
                $vCounter = 0
                $dstItemId = $null
                $lastVersionWhen = ""

                # Resolve createdDateTime once from source item metadata
                $srcCreated = $null
                if ($srcMeta) {
                    try {
                        if ($srcMeta.fileSystemInfo.createdDateTime) { $srcCreated = [datetime]$srcMeta.fileSystemInfo.createdDateTime }
                        elseif ($srcMeta.createdDateTime) { $srcCreated = [datetime]$srcMeta.createdDateTime }
                    } catch {}
                }

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
                    if ($vCounter -eq $versions.Count) {
                        $vBytes = Download-DriveFile -DriveId $SrcDriveId -ItemId $Item.Id -Token $SrcToken
                    } else {
                        $vBytes = Download-DriveFileVersion -DriveId $SrcDriveId -ItemId $Item.Id -VersionId $versionId -Token $SrcToken
                    }
                    Upload-BytesToDrivePath -DriveId $DstDriveId -RelPath $Item.RelPath -Bytes $vBytes -Token $DstToken
                    $actorForLog = if ($versionBy) { $versionBy } elseif ($versionActor) { $versionActor } else { "inconnu" }
                    Write-Log "DEBUG" "      Version $vCounter/$($versions.Count) rejouee (id=$versionId, par=$actorForLog, $($vBytes.Length) o)"

                    # Resolve dst item ID once after first upload
                    if (-not $dstItemId) {
                        $dstMeta = Get-DriveItemByRelPath -DriveId $DstDriveId -RelPath $Item.RelPath -Token $DstToken
                        if ($dstMeta -and $dstMeta.id) { $dstItemId = [string]$dstMeta.id }
                    }

                    if ($dstItemId) {
                        # Capture upload version ID before any PATCH/VVI changes it
                        $preOpVerId = $null
                        try {
                            $preOpVers = @(Get-DriveFileVersions -DriveId $DstDriveId -ItemId $dstItemId -Token $DstToken)
                            if ($preOpVers.Count -gt 0) { $preOpVerId = [string]($preOpVers | Select-Object -Last 1).id }
                        } catch {}

                        $vDate        = if ($versionWhen) { try { [datetime]$versionWhen } catch { $null } } else { $null }
                        $createdToUse = if ($srcCreated) { $srcCreated } elseif ($vDate) { $vDate } else { $null }
                        $srcVersionEmail = Normalize-Email -Value $versionBy
                        $mappedEmail = $null
                        if ($srcVersionEmail -and $script:UserMapping.ContainsKey($srcVersionEmail)) {
                            $mappedEmail = Normalize-Email -Value $script:UserMapping[$srcVersionEmail]
                        } elseif ($srcVersionEmail) {
                            $mappedEmail = $srcVersionEmail
                        }

                        $opDone = $false

                        # Primary: single VVI call setting Editor + Modified together — avoids any separate PATCH that would reset the author
                        if ($mappedEmail -and $versionWhen -and $script:DstSpToken -and $script:DstSiteInfo) {
                            try {
                                $serverRelPath = $script:DstSiteInfo.LibraryPath.TrimEnd('/') + '/' + ($Item.RelPath -replace '\\', '/')
                                try { Invoke-SharePointEnsureUser -SpToken $script:DstSpToken -HostUrl $script:DstSiteInfo.HostUrl -SitePath $script:DstSiteInfo.SitePath -Email $mappedEmail | Out-Null } catch {}
                                Set-SharePointEditor -SpToken $script:DstSpToken -HostUrl $script:DstSiteInfo.HostUrl `
                                    -SitePath $script:DstSiteInfo.SitePath -ServerRelPath $serverRelPath `
                                    -EditorEmail $mappedEmail -ModifiedDate $versionWhen
                                Write-Log "DEBUG" "      v$vCounter : auteur+date via VVI ($versionBy -> $mappedEmail / $versionWhen)"
                                $opDone = $true
                            } catch {
                                Write-Log "WARN" "      v$vCounter VVI echoue : $($_.Exception.Message)"
                            }
                        }

                        # Fallback: PATCH fileSystemInfo for date only (author stays as Application SharePoint)
                        if (-not $opDone -and $vDate -and $createdToUse) {
                            try {
                                Set-DriveItemFileTimestamps -DriveId $DstDriveId -ItemId $dstItemId -Token $DstToken -CreatedUtc $createdToUse -ModifiedUtc $vDate
                                Write-Log "DEBUG" "      v$vCounter : date via PATCH ($versionWhen)"
                                $opDone = $true
                            } catch { Write-Log "WARN" "      v$vCounter PATCH date echoue : $($_.Exception.Message)" }
                        }

                        # Delete the upload version if a new one was created by VVI or PATCH
                        if ($opDone -and $preOpVerId) {
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
                            Write-Log "WARN" "      v$vCounter : auteur non restaurable (email source introuvable). Acteur source='$actorForLog'"
                        }
                    }
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

                Write-Log "WARN" "Graph ne permet pas de forcer l'auteur des versions. Les versions cible seront attribuees au compte d'execution."
                Write-Log "SUCCESS" "✅ $($Item.RelPath) - historique copie ($($versions.Count) version(s))"
                return
            }
        }

        $bytes = Download-DriveFile -DriveId $SrcDriveId -ItemId $Item.Id -Token $SrcToken
        Write-Log "INFO" "   ⬆  $($Item.RelPath)"
        Upload-BytesToDrivePath -DriveId $DstDriveId -RelPath $Item.RelPath -Bytes $bytes -Token $DstToken

        if ($PreserveFileTimestamps -and $srcMeta) {
            try {
                $dstMeta = Get-DriveItemByRelPath -DriveId $DstDriveId -RelPath $Item.RelPath -Token $DstToken
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
                $dstMetaFields = Get-DriveItemByRelPath -DriveId $DstDriveId -RelPath $Item.RelPath -Token $DstToken
                if ($dstMetaFields -and $dstMetaFields.id) {
                    Copy-DriveItemMetadata -SrcDriveId $SrcDriveId -SrcItemId $Item.Id -DstDriveId $DstDriveId -DstItemId ([string]$dstMetaFields.id) -SrcToken $SrcToken -DstToken $DstToken -ItemRelPath $Item.RelPath
                }
            } catch {
                Write-Log "WARN" "Impossible de copier les metadonnees du fichier '$($Item.RelPath)' : $($_.Exception.Message)"
            }
        }

        if ($CopyVersionHistory) {
            Write-Log "WARN" "Graph ne permet pas de forcer l'auteur des versions. La version cible est attribuee au compte d'execution."
        }

        if ($CopyPermissions) {
            try {
                $dstMetaPerm = Get-DriveItemByRelPath -DriveId $DstDriveId -RelPath $Item.RelPath -Token $DstToken
                if ($dstMetaPerm -and $dstMetaPerm.id) {
                    Copy-DriveItemPermissions -SrcDriveId $SrcDriveId -SrcItemId $Item.Id -DstDriveId $DstDriveId -DstItemId ([string]$dstMetaPerm.id) -SrcToken $SrcToken -DstToken $DstToken -ItemRelPath $Item.RelPath
                }
            } catch {
                Write-Log "WARN" "Impossible de copier les permissions du fichier '$($Item.RelPath)' : $($_.Exception.Message)"
            }
        }

        Write-Log "SUCCESS" "✅ $($Item.RelPath) ($($bytes.Length) o)"

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

    # Ensure baseline mapping so versions modified by the owner can be restored without explicit CSV mapping line.
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
                Write-Log "INFO" "Metadonnees racine OneDrive traitees"
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

    # Dossiers
    Write-Log "INFO" "=== CREATION DES DOSSIERS ==="
    foreach ($folder in ($folders | Sort-Object { ($_.RelPath -split '/').Count }, RelPath)) {
        $dstToken = Get-AppToken -TenantId $app_cible.tenant_id -ClientId $app_cible.client_id.Trim() `
                    -Scope "https://graph.microsoft.com/.default" `
                    -ClientSecret $app_cible.client_secret `
                    -CertThumbprint $app_cible.certificate_thumbprint `
                    -CertPath $app_cible.certificate_path `
                    -CertPassword $app_cible.certificate_password
        Copy-DriveItem -Item $folder -SrcDriveId $srcDriveId -DstDriveId $dstDriveId `
                       -SrcToken $srcToken -DstToken $dstToken
    }

    # Fichiers
    Write-Log "INFO" "=== COPIE DES FICHIERS ==="
    $counter = 0
    foreach ($file in $files) {
        $counter++
        Write-Log "INFO" "[$counter/$($files.Count)] $($file.RelPath) ($($file.Size) o)"
        $srcToken = Get-GraphToken -TenantId $app_source.tenant_id `
                    -ClientId $app_source.client_id -ClientSecret $app_source.client_secret
        $dstToken = Get-AppToken -TenantId $app_cible.tenant_id -ClientId $app_cible.client_id.Trim() `
                    -Scope "https://graph.microsoft.com/.default" `
                    -ClientSecret $app_cible.client_secret `
                    -CertThumbprint $app_cible.certificate_thumbprint `
                    -CertPath $app_cible.certificate_path `
                    -CertPassword $app_cible.certificate_password
        Copy-DriveItem -Item $file -SrcDriveId $srcDriveId -DstDriveId $dstDriveId `
                       -SrcToken $srcToken -DstToken $dstToken
    }

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
    $migrations = Import-Csv -LiteralPath $CsvPath -Delimiter ';' -Encoding UTF8

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

        $errAvant = $script:ErrRows.Count

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