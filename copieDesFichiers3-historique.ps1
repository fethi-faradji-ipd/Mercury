
# Script PowerShell - Copie SharePoint inter-tenant (fichiers + permissions + versions best-effort)
# Version: 9.3


# =========================
# PARAMÈTRES
# =========================
$connexion_source = @{
        tenant_Id = "85a5a352-25d6-4894-a97d-221cd1712dd2"
        client_Id = "684f670e-6931-4120-b053-f8c9538744f6"
        client_Secret = "O5s8Q~Z~e_d1-dSQ.vt4uFTeR8c4Xssaa093Racz"
}

$connexion_cible = @{
    tenant_id     = "2eea08b8-1972-447b-ad43-d044d042500a"
    client_id     = "821109a4-6e0e-48e8-b477-8f9b70aa32a4"
    client_secret = "Yhy8Q~mBxTrYqx~ve7enqTQatjwhEZ5.34-q-a4j"
   # client_secret ="66A8Q~t2-Q_vZp54cDnf8TH~3kl0vp1UyHMZ8bVu"
}


# =========================
# PARAMÈTRES (hors secrets)
# =========================
$site_url_source = "https://ipdlab.sharepoint.com/sites/SERVICES_GENERAUX/"
$site_url_cible  = "https://infoprodigital365.sharepoint.com/sites/WH_GPA-test/"

$chemin_csv      = "E:\SharePoint_Permissions_Export_src9999.csv"    # Path,ItemType,GrantedTo,grantedToID,TargetType,Role
$translation_csv = "E:\UKREiiF_FinalCopy_path.csv"                # Source,Destination
$mapping_csv     = "E:\UserMapping.csv"                           # SourceId,SourceDisplayName,TargetType,DestinationDisplayName

$ForceOverwrite = $true
$ChunkSize      = 10MB
$CopierVersions = $true

# Ex: "Documents" ou "Drive ECHOLINE" si le CSV commence par ce préfixe
$CsvLibraryPrefixToStrip = ""

# =========================
# VARIABLES GLOBALES
# =========================
$script:RunId = (Get-Date -Format "yyyy-MM-dd_HHmmss")

$script:BasePath = if ($PSScriptRoot) {
    $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    (Get-Location).Path
}

if (-not (Test-Path -LiteralPath $script:BasePath)) {
    New-Item -ItemType Directory -Path $script:BasePath -Force | Out-Null
}

$script:LogPath    = Join-Path $script:BasePath ("\log\SharePoint_Copy_{0}.log" -f $script:RunId)
$script:ErrCsvPath = Join-Path $script:BasePath ("\log\SharePoint_Errors_{0}.csv" -f $script:RunId)
$script:ErrRows    = New-Object System.Collections.Generic.List[object]

$global:UrlTranslations      = @{}
$global:UserMapping          = @{}
$global:UserMappingByName    = @{}
$script:ResolvedIdCache      = @{}
$script:ResolvedSpGroupCache = @{}
$script:TokenCache           = @{}
$script:SrcAllDrives         = @()

Add-Type -AssemblyName System.Net.Http
$script:HttpClient = [System.Net.Http.HttpClient]::new()
$script:HttpClient.Timeout = [TimeSpan]::FromMinutes(30)

# =========================
# VALIDATION CONNEXIONS
# =========================
function Test-ConnexionSchema {
    param([hashtable]$Conn, [string]$Name)

    $required = @("tenant_id","client_id","client_secret")
    foreach ($k in $required) {
        if (-not $Conn.ContainsKey($k) -or [string]::IsNullOrWhiteSpace([string]$Conn[$k])) {
            throw "Connexion '$Name' invalide : clé '$k' manquante/vide."
        }
    }
}

# =========================
# LOGGING (REDACTION)
# =========================
function Protect-LogText([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $t = $Text

    $t = [regex]::Replace($t, '(?i)Bearer\s+[A-Za-z0-9\-\._~\+\/]+=*', 'Bearer ***')
    $t = [regex]::Replace($t, 'eyJ[A-Za-z0-9_\-]+\.[A-Za-z0-9_\-]+\.?[A-Za-z0-9_\-]*', '***JWT***')
    $t = [regex]::Replace($t, '(?i)(client[_-]?secret\s*[:=]\s*)[^,\s;]+', '$1***')
    $t = [regex]::Replace($t, '(?i)("client_secret"\s*:\s*")[^"]+(")', '$1***$2')
    $t = [regex]::Replace($t, '\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b', '***GUID***')
    $t = [regex]::Replace($t, '(?i)\b(driveid|itemid|siteid|objectid|principalid)\s*[:=]\s*[^,\s\)]+', '$1=***')
    $t = [regex]::Replace($t, '(?i)("objectId"\s*:\s*")[^"]+(")', '$1***$2')

    return $t
}

function Write-Log {
    param(
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG')][string]$Level = 'INFO',
        [string]$Message
    )

    $safe = Protect-LogText $Message
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[{0}] [{1}] {2}" -f $ts, $Level, $safe

    switch ($Level) {
        "SUCCESS" { Write-Host $line -ForegroundColor Green }
        "ERROR"   { Write-Host $line -ForegroundColor Red }
        "WARN"    { Write-Host $line -ForegroundColor Yellow }
        "INFO"    { Write-Host $line -ForegroundColor Cyan }
        "DEBUG"   { Write-Host $line -ForegroundColor Gray }
        default   { Write-Host $line }
    }

    Add-Content -Path $script:LogPath -Value $line -Encoding UTF8
}

function Append-Err {
    param([string]$Rel, [string]$ErrMsg, [string]$Phase = "")
    $script:ErrRows.Add([pscustomobject]@{
        horodatage     = Get-Date
        chemin_relatif = $Rel
        phase          = $Phase
        erreur         = (Protect-LogText $ErrMsg)
    }) | Out-Null
}

function Get-ErrorDetail([object]$err) {
    $msg = $err.Exception.Message
    if ($err.ErrorDetails -and $err.ErrorDetails.Message) {
        $msg = "$msg | $($err.ErrorDetails.Message)"
    }
    return $msg
}

# =========================
# OUTILS
# =========================
function Normalize-TextC([string]$s) {
    if ([string]::IsNullOrEmpty($s)) { return $s }
    return $s.Normalize([Text.NormalizationForm]::FormC)
}

function Normalize-RelPath([string]$rel) {
    if (-not $rel) { return "" }
    $r = ($rel.Trim() -replace '\\','/').TrimStart('/').TrimEnd('/ ')
    return (Normalize-TextC $r)
}

function Strip-LibraryPrefix([string]$rel) {
    if ([string]::IsNullOrWhiteSpace($CsvLibraryPrefixToStrip)) { return $rel }
    $prefix = $CsvLibraryPrefixToStrip.Trim().TrimEnd('/') + "/"

    if ($rel.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $rel.Substring($prefix.Length)
    }
    if ($rel -ieq $CsvLibraryPrefixToStrip.Trim()) { return "" }
    return $rel
}

function UrlEncode-Segment([string]$seg) {
    if ([string]::IsNullOrWhiteSpace($seg)) { return "" }
    [System.Uri]::EscapeDataString((Normalize-TextC $seg))
}

function Encode-RelForGraph([string]$rel) {
    $segs = ($rel -replace '\\','/') -split '/' | Where-Object { $_ -ne '' }
    return [string]::Join('/', ($segs | ForEach-Object { UrlEncode-Segment $_ }))
}

function Build-AuthHeader([string]$token) { @{ Authorization = "Bearer $token" } }

function Get-HostAndPathFromSiteUrl([string]$siteUrl) {
    $u = [Uri]$siteUrl
    $p = $u.AbsolutePath
    if ($p.Length -gt 1) { $p = $p.TrimEnd('/') } # fix 400 Graph
    [pscustomobject]@{ host = $u.Host; path = $p }
}

function Get-RetryDelayMs($attempt, $retryAfter) {
    if ($retryAfter) {
        $s = 0
        if ([int]::TryParse($retryAfter, [ref]$s)) { return [int]($s * 1000) }
    }
    $base = [math]::Pow(2, [math]::Max(0, $attempt - 1)) * 500
    $jitter = Get-Random -Min 0 -Max 300
    return [int]([math]::Min(32000, $base + $jitter))
}

function Is-TokenErrorText([string]$txt) {
    $txt -match '401|invalid[_-]?token|expired|invalid[_-]?grant|AADSTS'
}

# =========================
# TOKENS
# =========================
function Get-AccessTokenResource {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret,
        [Parameter(Mandatory)][string]$ResourceAppIdUri
    )

    $cacheKey = "$TenantId|$ClientId|$ResourceAppIdUri"
    $now = Get-Date
    $skew = [TimeSpan]::FromMinutes(5)

    if ($script:TokenCache.ContainsKey($cacheKey)) {
        $e = $script:TokenCache[$cacheKey]
        if ($e.expires -gt $now.Add($skew)) { return $e.token }
    }

    $uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
        scope         = "$ResourceAppIdUri/.default"
    }

    $resp = Invoke-RestMethod -Method POST -Uri $uri -Body $body -ContentType "application/x-www-form-urlencoded"
    $token = $resp.access_token
    $exp = (Get-Date).AddSeconds([int]$resp.expires_in)
    $script:TokenCache[$cacheKey] = @{ token = $token; expires = $exp }
    return $token
}

function Get-AccessToken {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret
    )
    Get-AccessTokenResource -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -ResourceAppIdUri "https://graph.microsoft.com"
}

function Get-SharePointAccessToken {
    param(
        [Parameter(Mandatory)][hashtable]$Connexion,
        [Parameter(Mandatory)][string]$SiteUrl
    )
    $authority = ([Uri]$SiteUrl).GetLeftPart([System.UriPartial]::Authority)
    Get-AccessTokenResource -TenantId $Connexion.tenant_id -ClientId $Connexion.client_id -ClientSecret $Connexion.client_secret -ResourceAppIdUri $authority
}

# =========================
# GRAPH AVEC RETRY
# =========================
function Invoke-GraphJson {
    param(
        [string]$Method = "GET",
        [string]$Uri,
        [hashtable]$Connexion,
        [ref]$TokenRef,
        $Body = $null,
        [string]$ContentType = "application/json",
        [int]$MaxRetries = 5
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        $TokenRef.Value = Get-AccessToken -TenantId $Connexion.tenant_id -ClientId $Connexion.client_id -ClientSecret $Connexion.client_secret
        try {
            $headers = Build-AuthHeader $TokenRef.Value
            $payload = $Body
            if ($null -ne $payload -and $ContentType -eq "application/json" -and ($payload -isnot [string])) {
                $payload = $payload | ConvertTo-Json -Depth 12
            }
            return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $payload -ContentType $ContentType
        } catch {
            $status = $null; $retryAfter = $null
            if ($_.Exception.Response) {
                try { $status = [int]$_.Exception.Response.StatusCode } catch {}
                if ($_.Exception.Response.Headers["Retry-After"]) { $retryAfter = $_.Exception.Response.Headers["Retry-After"] }
            }

            $msg = $_.Exception.Message
            $shouldRetry = ($status -eq 429) -or ($status -eq 408) -or ($status -ge 500 -and $status -lt 600) -or (Is-TokenErrorText $msg)

            if ($shouldRetry -and $attempt -lt $MaxRetries) {
                $delay = Get-RetryDelayMs $attempt $retryAfter
                Write-Log "WARN" ("Retry {0} ({1}) Graph. Attente {2}ms" -f $attempt, $status, $delay)
                Start-Sleep -Milliseconds $delay
                continue
            }
            throw
        }
    }
}

# =========================
# SITE / DRIVE
# =========================
function Get-SiteAndDrive {
    param([string]$SiteUrl, [string]$AccessToken, [string[]]$PreferredLibraries = @("Documents","Dokument"))

    $hp = Get-HostAndPathFromSiteUrl $SiteUrl
    $headers = Build-AuthHeader $AccessToken

    $siteUri = "https://graph.microsoft.com/v1.0/sites/$($hp.host):$($hp.path)?`$select=id,webUrl,displayName"
    Write-Log "DEBUG" ("Resolve site URI: {0}" -f $siteUri)

    $site = Invoke-RestMethod -Method GET -Uri $siteUri -Headers $headers
    $drives = Invoke-RestMethod -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives?`$select=id,name,driveType" -Headers $headers

    $libs = $drives.value | Where-Object { $_.driveType -eq "documentLibrary" }
    $drive = $null
    foreach ($p in $PreferredLibraries) {
        $drive = $libs | Where-Object { $_.name -ieq $p } | Select-Object -First 1
        if ($drive) { break }
    }
    if (-not $drive) { $drive = $libs | Select-Object -First 1 }

    [pscustomobject]@{ Site = $site; Drive = $drive; AllDrives = $libs }
}

function Get-AllDrivesFromSite {
    param([string]$SiteId, [string]$AccessToken)
    $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives?`$select=id,name,driveType"
    $r = Invoke-RestMethod -Method GET -Uri $uri -Headers (Build-AuthHeader $AccessToken)
    $r.value | Where-Object { $_.driveType -eq "documentLibrary" }
}

function Item-Exists {
    param([string]$DriveId, [string]$Rel, [string]$AccessToken)

    $encRel = Encode-RelForGraph $Rel
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$encRel"
    try {
        $it = Invoke-RestMethod -Method GET -Uri $uri -Headers (Build-AuthHeader $AccessToken)
        [pscustomobject]@{
            Exists   = $true
            IsFolder = ($null -ne $it.folder)
            Id       = $it.id
            Size     = $it.size
            WebUrl   = $it.webUrl
        }
    } catch {
        if ($_.Exception.Response -and [int]$_.Exception.Response.StatusCode -eq 404) {
            [pscustomobject]@{ Exists = $false; IsFolder = $false; Id = $null; Size = 0; WebUrl = $null }
        } else { throw }
    }
}

function Ensure-FolderPath {
    param([string]$DriveId, [string]$FolderRel, [string]$AccessToken)

    $rel = Normalize-RelPath $FolderRel
    $segs = $rel -split '/' | Where-Object { $_ -ne '' }
    if ($segs.Count -eq 0) { return $null }

    $headers = Build-AuthHeader $AccessToken
    $parentId = "root"
    $lastId = $null

    foreach ($segRaw in $segs) {
        $seg = Normalize-TextC $segRaw
        $enc = UrlEncode-Segment $seg

        $checkUri = if ($parentId -eq "root") {
            "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$enc"
        } else {
            "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$parentId`:/$enc"
        }

        $child = $null
        try {
            $child = Invoke-RestMethod -Method GET -Uri $checkUri -Headers $headers
        } catch {
            $st = $null
            if ($_.Exception.Response) { try { $st = [int]$_.Exception.Response.StatusCode } catch {} }
            if ($st -ne 404) { throw }

            $createUri = if ($parentId -eq "root") {
                "https://graph.microsoft.com/v1.0/drives/$DriveId/root/children"
            } else {
                "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$parentId/children"
            }

            $body = @{
                name = $seg
                folder = @{}
                "@microsoft.graph.conflictBehavior" = "replace"
            } | ConvertTo-Json -Depth 4

            $child = Invoke-RestMethod -Method POST -Uri $createUri -Headers $headers -Body $body -ContentType "application/json"
        }

        $parentId = $child.id
        $lastId = $child.id
    }

    return $lastId
}

# =========================
# DOWNLOAD / UPLOAD
# =========================
function Download-FileContent {
    param(
        [string]$DriveId,
        [string]$RelativePath,
        [hashtable]$Connexion,
        [ref]$TokenRef,
        [int]$MaxRetries = 5
    )

    $encRel = Encode-RelForGraph $RelativePath
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$encRel`:/content"

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        $TokenRef.Value = Get-AccessToken -TenantId $Connexion.tenant_id -ClientId $Connexion.client_id -ClientSecret $Connexion.client_secret
        $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $uri)
        $req.Headers.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $TokenRef.Value)

        $resp = $script:HttpClient.SendAsync($req).Result
        if ($resp.IsSuccessStatusCode) { return $resp.Content.ReadAsByteArrayAsync().Result }

        $status = [int]$resp.StatusCode
        $retryAfter = $null
        if ($resp.Headers.RetryAfter -and $resp.Headers.RetryAfter.Delta) {
            $retryAfter = [int][math]::Ceiling($resp.Headers.RetryAfter.Delta.TotalSeconds)
        }

        if (($status -in @(401,403,429) -or ($status -ge 500 -and $status -lt 600)) -and $attempt -lt $MaxRetries) {
            $delay = Get-RetryDelayMs $attempt $retryAfter
            Write-Log "WARN" ("Download retry {0} ({1}). Attente {2}ms" -f $attempt, $status, $delay)
            Start-Sleep -Milliseconds $delay
            continue
        }

        throw "Download HTTP $status"
    }
}

function Download-FileVersionContent {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$VersionId,
        [hashtable]$Connexion,
        [ref]$TokenRef,
        [int]$MaxRetries = 5
    )

    $versionIdEnc = [System.Uri]::EscapeDataString($VersionId)
    $contentUri   = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ItemId/versions/$versionIdEnc/content"
    $metaUri      = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ItemId/versions/$versionIdEnc"

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        $TokenRef.Value = Get-AccessToken -TenantId $Connexion.tenant_id -ClientId $Connexion.client_id -ClientSecret $Connexion.client_secret

        # 1) tentative directe /content
        try {
            $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $contentUri)
            $req.Headers.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $TokenRef.Value)
            $resp = $script:HttpClient.SendAsync($req).Result
            if ($resp.IsSuccessStatusCode) {
                return $resp.Content.ReadAsByteArrayAsync().Result
            }

            $status = [int]$resp.StatusCode
            if ($status -notin @(400,404)) {
                if (($status -in @(401,403,429) -or ($status -ge 500 -and $status -lt 600)) -and $attempt -lt $MaxRetries) {
                    Start-Sleep -Milliseconds (Get-RetryDelayMs $attempt $null)
                    continue
                }
            }
        } catch {}

        # 2) fallback downloadUrl
        try {
            $ver = Invoke-RestMethod -Method GET -Uri $metaUri -Headers (Build-AuthHeader $TokenRef.Value)
            $dl = $ver.'@microsoft.graph.downloadUrl'
            if (-not [string]::IsNullOrWhiteSpace($dl)) {
                $req2 = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $dl)
                $resp2 = $script:HttpClient.SendAsync($req2).Result
                if ($resp2.IsSuccessStatusCode) {
                    return $resp2.Content.ReadAsByteArrayAsync().Result
                }
            }
        } catch {}

        if ($attempt -lt $MaxRetries) {
            Start-Sleep -Milliseconds (Get-RetryDelayMs $attempt $null)
            continue
        }

        throw "Version download impossible pour version '$VersionId' (content + downloadUrl en échec)"
    }
}

function Upload-SmallFile {
    param(
        [string]$DriveId,
        [string]$RelativePath,
        [byte[]]$Bytes,
        [hashtable]$Connexion,
        [ref]$TokenRef
    )

    $rel = Normalize-RelPath $RelativePath
    $parent = ([System.IO.Path]::GetDirectoryName($rel) -replace '\\','/')
    if ($parent -and $parent -ne ".") {
        $TokenRef.Value = Get-AccessToken -TenantId $Connexion.tenant_id -ClientId $Connexion.client_id -ClientSecret $Connexion.client_secret
        Ensure-FolderPath -DriveId $DriveId -FolderRel $parent -AccessToken $TokenRef.Value | Out-Null
    }

    $encRel = Encode-RelForGraph $rel
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$encRel`:/content"

    $TokenRef.Value = Get-AccessToken -TenantId $Connexion.tenant_id -ClientId $Connexion.client_id -ClientSecret $Connexion.client_secret
    Invoke-RestMethod -Method PUT -Uri $uri -Headers (Build-AuthHeader $TokenRef.Value) -Body $Bytes -ContentType "application/octet-stream"
}

function Upload-LargeFile {
    param(
        [string]$DriveId,
        [string]$RelativePath,
        [byte[]]$Bytes,
        [hashtable]$Connexion,
        [ref]$TokenRef,
        [int]$ChunkBytes = 8388608
    )

    $rel = Normalize-RelPath $RelativePath
    $parent = ([System.IO.Path]::GetDirectoryName($rel) -replace '\\','/')
    if ($parent -and $parent -ne ".") {
        $TokenRef.Value = Get-AccessToken -TenantId $Connexion.tenant_id -ClientId $Connexion.client_id -ClientSecret $Connexion.client_secret
        Ensure-FolderPath -DriveId $DriveId -FolderRel $parent -AccessToken $TokenRef.Value | Out-Null
    }

    $encRel = Encode-RelForGraph $rel
    $createUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$encRel`:/createUploadSession"
    $session = Invoke-GraphJson -Method POST -Uri $createUri -Connexion $Connexion -TokenRef $TokenRef -Body @{
        item = @{ "@microsoft.graph.conflictBehavior" = "replace" }
    }

    $uploadUrl = $session.uploadUrl
    $total = $Bytes.Length
    $pos = 0
    $att = 0
    $lastResponse = $null

    while ($pos -lt $total) {
        $att++
        $chunk = [Math]::Min($ChunkBytes, $total - $pos)
        $from = $pos
        $to = $pos + $chunk - 1

        $content = [System.Net.Http.ByteArrayContent]::new($Bytes, $from, $chunk)
        $content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::new("application/octet-stream")
        $null = $content.Headers.TryAddWithoutValidation("Content-Range", ("bytes {0}-{1}/{2}" -f $from, $to, $total))

        $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Put, $uploadUrl)
        $req.Content = $content
        $resp = $script:HttpClient.SendAsync($req).Result

        if ($resp.IsSuccessStatusCode) {
            $txt = $resp.Content.ReadAsStringAsync().Result
            if ($txt) { try { $lastResponse = $txt | ConvertFrom-Json } catch {} }
            $pos += $chunk
            continue
        }

        $status = [int]$resp.StatusCode
        $retryAfter = $null
        if ($resp.Headers.RetryAfter -and $resp.Headers.RetryAfter.Delta) {
            $retryAfter = [int][math]::Ceiling($resp.Headers.RetryAfter.Delta.TotalSeconds)
        }

        if (($status -eq 429 -or ($status -ge 500 -and $status -lt 600)) -and $pos -lt $total) {
            $delay = Get-RetryDelayMs $att $retryAfter
            Write-Log "WARN" ("Chunk retry ({0}). Attente {1}ms" -f $status, $delay)
            Start-Sleep -Milliseconds $delay
            continue
        }

        $errTxt = $resp.Content.ReadAsStringAsync().Result
        throw "Upload HTTP $status : $errTxt"
    }

    return $lastResponse
}

# =========================
# TRADUCTIONS + MAPPING
# =========================
function Load-UrlTranslations {
    param([string]$TranslationFile)

    try {
        Write-Log "INFO" ("Chargement traductions: {0}" -f $TranslationFile)

        if (-not (Test-Path $TranslationFile)) {
            Write-Log "WARN" "Fichier de traduction introuvable"
            return
        }

        $csv = Import-Csv $TranslationFile
        $n = 0

        foreach ($r in $csv) {
            $src = if ($r.PSObject.Properties['Source']) { [string]$r.Source } else { "" }
            $dst = if ($r.PSObject.Properties['Destination']) { [string]$r.Destination } else { "" }

            $src = (Normalize-RelPath ($src.Trim().Trim('"')))
            $dst = (Normalize-RelPath ($dst.Trim().Trim('"')))

            if ([string]::IsNullOrWhiteSpace($src) -or [string]::IsNullOrWhiteSpace($dst)) { continue }

            $global:UrlTranslations[$src] = $dst
            $n++
        }

        Write-Log "SUCCESS" ("Traductions chargées: {0}" -f $n)
    } catch {
        Write-Log "ERROR" ("Erreur chargement traductions: {0}" -f $_.Exception.Message)
    }
}

function Apply-UrlTranslations {
    param([string]$OriginalUrl)

    if ([string]::IsNullOrWhiteSpace($OriginalUrl)) { return $OriginalUrl }
    if ($global:UrlTranslations.Count -eq 0) { return $OriginalUrl }

    $t = $OriginalUrl
    foreach ($kv in $global:UrlTranslations.GetEnumerator()) {
        if ([string]::IsNullOrWhiteSpace($kv.Key)) { continue }
        if ($t.Contains($kv.Key)) { $t = $t.Replace($kv.Key, $kv.Value) }
    }
    return $t
}

function Generate-UserMappingTemplate {
    param([object[]]$Rows, [string]$MappingFile)

    $unique = $Rows | Select-Object grantedToID, GrantedTo, TargetType -Unique
    $template = $unique | ForEach-Object {
        [pscustomobject]@{
            SourceId               = $_.grantedToID
            SourceDisplayName      = $_.GrantedTo
            TargetType             = $_.TargetType
            DestinationDisplayName = ""
        }
    }

    $template | Export-Csv -Path $MappingFile -NoTypeInformation -Encoding UTF8
    Write-Log "WARN" ("Template mapping généré: {0}" -f $MappingFile)
}

function Load-UserMapping {
    param([string]$MappingFile)

    if (-not (Test-Path $MappingFile)) {
        Write-Log "WARN" ("Mapping introuvable: {0}" -f $MappingFile)
        return
    }

    $csv = Import-Csv $MappingFile -Encoding UTF8
    $idCount = 0; $nameCount = 0

    foreach ($row in $csv) {
        $srcId   = if ($row.PSObject.Properties['SourceId'])               { $row.SourceId.Trim() }               else { '' }
        $srcName = if ($row.PSObject.Properties['SourceDisplayName'])      { $row.SourceDisplayName.Trim() }      else { '' }
        $dstName = if ($row.PSObject.Properties['DestinationDisplayName']) { $row.DestinationDisplayName.Trim() } else { '' }
        $tt      = if ($row.PSObject.Properties['TargetType'])             { $row.TargetType.Trim() }             else { '' }

        if (-not $dstName) { continue }

        $entry = @{ DisplayName = $dstName; TargetType = $tt }
        if ($srcId)   { $global:UserMapping[$srcId] = $entry; $idCount++ }
        if ($srcName) { $global:UserMappingByName[$srcName] = $entry; $nameCount++ }
    }

    Write-Log "SUCCESS" ("Mappages chargés: ID={0}, Name={1}" -f $idCount, $nameCount)
}

function Resolve-DestinationObjectId {
    param([string]$DisplayName, [string]$TargetType, [ref]$DstTokenRef)

    $type = $TargetType.Trim().ToLowerInvariant()
    $isGroup = $type -in @('group','securitygroup','m365group','microsoft365group','aadgroup')
    $cacheKey = "$type|$DisplayName"

    if ($script:ResolvedIdCache.ContainsKey($cacheKey)) { return $script:ResolvedIdCache[$cacheKey] }

    $safe = $DisplayName.Replace("'","''")
    $uri = if ($isGroup) {
        "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe'&`$select=id&`$top=1"
    } else {
        "https://graph.microsoft.com/v1.0/users?`$filter=displayName eq '$safe'&`$select=id&`$top=1"
    }

    try {
        $r = Invoke-GraphJson -Method GET -Uri $uri -Connexion $connexion_cible -TokenRef $DstTokenRef
        $obj = $r.value | Select-Object -First 1
        $id = if ($obj) { $obj.id } else { $null }
        $script:ResolvedIdCache[$cacheKey] = $id
        return $id
    } catch {
        $script:ResolvedIdCache[$cacheKey] = $null
        return $null
    }
}

# =========================
# GROUPES SHAREPOINT (REST)
# =========================
function Get-SPFormDigest {
    param([string]$SiteUrl, [string]$AccessToken)

    $uri = $SiteUrl.TrimEnd('/') + "/_api/contextinfo"
    $headers = @{ Authorization = "Bearer $AccessToken"; Accept = "application/json;odata=nometadata" }
    (Invoke-RestMethod -Method POST -Uri $uri -Headers $headers).FormDigestValue
}

function Resolve-SharePointGroupId {
    param([string]$SiteUrl, [string]$GroupName, [string]$AccessToken)

    $cacheKey = "$($SiteUrl.ToLower())|$($GroupName.ToLower())"
    if ($script:ResolvedSpGroupCache.ContainsKey($cacheKey)) { return $script:ResolvedSpGroupCache[$cacheKey] }

    $safe = $GroupName.Replace("'","''")
    $uri = $SiteUrl.TrimEnd('/') + "/_api/web/sitegroups/getbyname('$safe')?`$select=Id"
    $headers = @{ Authorization = "Bearer $AccessToken"; Accept = "application/json;odata=nometadata" }

    try {
        $g = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
        $id = [int]$g.Id
        $script:ResolvedSpGroupCache[$cacheKey] = $id
        return $id
    } catch {
        $script:ResolvedSpGroupCache[$cacheKey] = $null
        return $null
    }
}

function Get-SPRoleDefId([string]$Role) {
    switch ($Role.Trim().ToLowerInvariant()) {
        "owner" { 1073741829 } # Full Control
        "write" { 1073741827 } # Contribute
        default { 1073741826 } # Read
    }
}

function Get-SPItemApiUrl {
    param([string]$SiteUrl, [string]$ServerRelativeUrl, [bool]$IsFolder)

    $safe = $ServerRelativeUrl.Replace("'","''")
    if ($IsFolder) {
        return $SiteUrl.TrimEnd('/') + "/_api/web/GetFolderByServerRelativePath(decodedurl='$safe')/ListItemAllFields"
    }
    return $SiteUrl.TrimEnd('/') + "/_api/web/GetFileByServerRelativePath(decodedurl='$safe')/ListItemAllFields"
}

function Ensure-UniquePermissions {
    param([string]$ItemApiUrl, [string]$AccessToken, [string]$Digest)

    $uri = "$ItemApiUrl/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)"
    $headers = @{
        Authorization     = "Bearer $AccessToken"
        Accept            = "application/json;odata=nometadata"
        "X-RequestDigest" = $Digest
    }
    $null = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers
}

function Grant-SPPermission {
    param([string]$ItemApiUrl, [int]$PrincipalId, [int]$RoleDefId, [string]$AccessToken, [string]$Digest)

    $uri = "$ItemApiUrl/roleassignments/addroleassignment(principalid=$PrincipalId,roledefid=$RoleDefId)"
    $headers = @{
        Authorization     = "Bearer $AccessToken"
        Accept            = "application/json;odata=nometadata"
        "X-RequestDigest" = $Digest
    }
    $null = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers
}

function Test-IsSpGroup([object]$PermRow) {
    $tt = if ($PermRow.PSObject.Properties['TargetType']) { $PermRow.TargetType.Trim().ToLowerInvariant() } else { "" }
    $dn = if ($PermRow.PSObject.Properties['GrantedTo'])   { $PermRow.GrantedTo.Trim() } else { "" }
    return ($tt -match 'sharepoint') -or ($dn -match '(Propriétaires|Membres|Visiteurs|Owners|Members|Visitors)')
}

# =========================
# VERSIONS
# =========================
function Get-DriveItemVersions {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$AccessToken
    )

    $all = New-Object System.Collections.Generic.List[object]
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ItemId/versions?`$select=id,lastModifiedDateTime,publication,size"

    while ($uri) {
        $r = Invoke-RestMethod -Method GET -Uri $uri -Headers (Build-AuthHeader $AccessToken)
        if ($r.value) {
            foreach ($v in $r.value) { $all.Add($v) | Out-Null }
        }
        $uri = $null
        if ($r.PSObject.Properties['@odata.nextLink'] -and $r.'@odata.nextLink') {
            $uri = $r.'@odata.nextLink'
        }
    }

    return @($all | Sort-Object { [datetime]$_.lastModifiedDateTime })
}

function Copy-FileWithVersions {
    param(
        [string]$Rel,
        [string]$SrcDriveId,
        [string]$SrcToken,
        [string]$DstDriveId,
        [string]$DstToken,
        [hashtable]$SrcConnexion,
        [hashtable]$DstConnexion
    )

    $originalRel   = Normalize-RelPath (Strip-LibraryPrefix (Normalize-RelPath $Rel))
    $translatedRel = Normalize-RelPath (Apply-UrlTranslations -OriginalUrl $originalRel)

    $srcItem = Item-Exists -DriveId $SrcDriveId -Rel $originalRel -AccessToken $SrcToken
    if (-not $srcItem.Exists -or $srcItem.IsFolder) {
        return @{ Ok = $false; ItemId = $null; TranslatedRel = $translatedRel }
    }

    $dstTokRef = [ref]$DstToken
    $srcTokRef = [ref]$SrcToken
    $lastUpload = $null

    $versions = @()
    if ($CopierVersions) {
        try {
            $versions = @(Get-DriveItemVersions -DriveId $SrcDriveId -ItemId $srcItem.Id -AccessToken $SrcToken)
            Write-Log "INFO" ("Versions source trouvées: {0}" -f $versions.Count)
        } catch {
            Write-Log "WARN" "Lecture des versions impossible, fallback copie simple"
            $versions = @()
        }
    }

    $historicVersions = @()
    if ($versions.Count -gt 1) {
        $historicVersions = @($versions | Select-Object -First ($versions.Count - 1))
    } elseif ($versions.Count -eq 1) {
        Write-Log "INFO" "Une seule version détectée (1.0 probable) : upload courant uniquement."
    }

    foreach ($ver in $historicVersions) {
        try {
            $comment = ""
            if ($ver.publication -and $ver.publication.comment) { $comment = [string]$ver.publication.comment }
            if ($comment) {
                Write-Log "INFO" ("Copie version {0} | commentaire: {1}" -f $ver.id, $comment)
            } else {
                Write-Log "INFO" ("Copie version {0}" -f $ver.id)
            }

            $bytes = Download-FileVersionContent -DriveId $SrcDriveId -ItemId $srcItem.Id -VersionId $ver.id -Connexion $SrcConnexion -TokenRef $srcTokRef

            $lastUpload = if ($bytes.Length -le 4MB) {
                Upload-SmallFile -DriveId $DstDriveId -RelativePath $translatedRel -Bytes $bytes -Connexion $DstConnexion -TokenRef $dstTokRef
            } else {
                Upload-LargeFile -DriveId $DstDriveId -RelativePath $translatedRel -Bytes $bytes -Connexion $DstConnexion -TokenRef $dstTokRef -ChunkBytes ([int]$ChunkSize)
            }
        } catch {
            $detail = Get-ErrorDetail $_
            Append-Err -Rel $originalRel -ErrMsg $detail -Phase "version-copy"
            Write-Log "WARN" ("Version ignorée ({0}) : {1}" -f $ver.id, $detail)
        }
    }

    $currentBytes = Download-FileContent -DriveId $SrcDriveId -RelativePath $originalRel -Connexion $SrcConnexion -TokenRef $srcTokRef
    $lastUpload = if ($currentBytes.Length -le 4MB) {
        Upload-SmallFile -DriveId $DstDriveId -RelativePath $translatedRel -Bytes $currentBytes -Connexion $DstConnexion -TokenRef $dstTokRef
    } else {
        Upload-LargeFile -DriveId $DstDriveId -RelativePath $translatedRel -Bytes $currentBytes -Connexion $DstConnexion -TokenRef $dstTokRef -ChunkBytes ([int]$ChunkSize)
    }

    $itemId = if ($lastUpload -and $lastUpload.id) { $lastUpload.id } else { $null }
    if (-not $itemId) {
        $check = Item-Exists -DriveId $DstDriveId -Rel $translatedRel -AccessToken $DstToken
        if ($check.Exists) { $itemId = $check.Id }
    }

    return @{ Ok = $true; ItemId = $itemId; TranslatedRel = $translatedRel }
}

# =========================
# PERMISSIONS
# =========================
function Apply-Item-Permissions {
    param(
        [string]$TranslatedRel,
        [string]$DstDriveId,
        [string]$DstSiteUrl,
        [ref]$DstTokenRef,
        [object[]]$PermRows,
        [string]$KnownItemId = ""
    )

    if (-not $PermRows -or $PermRows.Count -eq 0) { return }

    $TranslatedRel = Normalize-RelPath $TranslatedRel
    $itemId = $KnownItemId
    $webUrl = $null
    $isFolder = $false

    if (-not $itemId) {
        $encRel = Encode-RelForGraph $TranslatedRel
        $itemUri = "https://graph.microsoft.com/v1.0/drives/$DstDriveId/root:/$encRel?`$select=id,webUrl,folder,file"

        for ($i = 1; $i -le 8; $i++) {
            try {
                $item = Invoke-GraphJson -Method GET -Uri $itemUri -Connexion $connexion_cible -TokenRef $DstTokenRef
                $itemId = $item.id
                $webUrl = $item.webUrl
                $isFolder = ($null -ne $item.folder)
                break
            } catch {
                $st = $null
                if ($_.Exception.Response) { try { $st = [int]$_.Exception.Response.StatusCode } catch {} }
                if ($st -eq 404 -and $i -lt 8) { Start-Sleep -Seconds 3; continue }
                Write-Log "WARN" ("Impossible de récupérer item pour droits: {0}" -f $TranslatedRel)
                return
            }
        }
    }

    if (-not $itemId) { return }

    if (-not $webUrl) {
        try {
            $d = Invoke-GraphJson -Method GET -Uri "https://graph.microsoft.com/v1.0/drives/$DstDriveId/items/$itemId?`$select=webUrl,folder,file" -Connexion $connexion_cible -TokenRef $DstTokenRef
            $webUrl = $d.webUrl
            $isFolder = ($null -ne $d.folder)
        } catch {}
    }

    $serverRel = if ($webUrl) { ([Uri]$webUrl).AbsolutePath } else { $null }

    $spToken = $null
    $spDigest = $null
    $spItemApiUrl = $null
    $spUniqueBroken = $false

    foreach ($perm in $PermRows) {
        $srcId   = if ($perm.PSObject.Properties['grantedToID']) { $perm.grantedToID.Trim() } else { "" }
        $srcName = if ($perm.PSObject.Properties['GrantedTo'])   { $perm.GrantedTo.Trim() } else { "" }
        $srcRole = if ($perm.PSObject.Properties['Role'])        { $perm.Role.Trim() } else { "read" }

        $entry = $null
        if ($srcId -and $srcId -ne "Unknown" -and $global:UserMapping.ContainsKey($srcId)) {
            $entry = $global:UserMapping[$srcId]
        } elseif ($srcName -and $global:UserMappingByName.ContainsKey($srcName)) {
            $entry = $global:UserMappingByName[$srcName]
        }

        if (Test-IsSpGroup -PermRow $perm) {
            $targetGroupName = if ($entry) { $entry.DisplayName } else { $srcName }
            if (-not $targetGroupName -or -not $serverRel) { continue }

            try {
                if (-not $spToken) {
                    $spToken = Get-SharePointAccessToken -Connexion $connexion_cible -SiteUrl $DstSiteUrl
                    $spDigest = Get-SPFormDigest -SiteUrl $DstSiteUrl -AccessToken $spToken
                    $spItemApiUrl = Get-SPItemApiUrl -SiteUrl $DstSiteUrl -ServerRelativeUrl $serverRel -IsFolder:$isFolder
                }

                $spGroupId = Resolve-SharePointGroupId -SiteUrl $DstSiteUrl -GroupName $targetGroupName -AccessToken $spToken
                if (-not $spGroupId) {
                    Write-Log "WARN" "Groupe SP introuvable sur destination (permission ignorée)"
                    continue
                }

                if (-not $spUniqueBroken) {
                    Ensure-UniquePermissions -ItemApiUrl $spItemApiUrl -AccessToken $spToken -Digest $spDigest
                    $spUniqueBroken = $true
                }

                $roleDefId = Get-SPRoleDefId -Role $srcRole
                Grant-SPPermission -ItemApiUrl $spItemApiUrl -PrincipalId $spGroupId -RoleDefId $roleDefId -AccessToken $spToken -Digest $spDigest
                Write-Log "SUCCESS" ("Permission SP appliquée [{0}] -> {1}" -f $srcRole, $TranslatedRel)
            } catch {
                $detail = Get-ErrorDetail $_
                Append-Err -Rel $TranslatedRel -ErrMsg $detail -Phase "permissions-sp"
                Write-Log "WARN" "Erreur permission SP"
            }
            continue
        }

        if (-not $entry) {
            Write-Log "WARN" ("Pas de mapping (permission ignorée) -> {0}" -f $TranslatedRel)
            continue
        }

        $dstId = Resolve-DestinationObjectId -DisplayName $entry.DisplayName -TargetType $entry.TargetType -DstTokenRef $DstTokenRef
        if (-not $dstId) {
            Write-Log "WARN" ("Résolution Entra impossible (permission ignorée) -> {0}" -f $TranslatedRel)
            continue
        }

        $role = switch ($srcRole.ToLowerInvariant()) {
            "owner" { "write" }
            "write" { "write" }
            default { "read" }
        }

        $inviteUri = "https://graph.microsoft.com/v1.0/drives/$DstDriveId/items/$itemId/invite"
        $body = @{
            requireSignIn  = $true
            sendInvitation = $false
            roles          = @($role)
            recipients     = @(@{ objectId = $dstId })
        }

        try {
            $null = Invoke-GraphJson -Method POST -Uri $inviteUri -Connexion $connexion_cible -TokenRef $DstTokenRef -Body $body
            Write-Log "SUCCESS" ("Permission Entra appliquée [{0}] -> {1}" -f $role, $TranslatedRel)
        } catch {
            $detail = Get-ErrorDetail $_
            Append-Err -Rel $TranslatedRel -ErrMsg $detail -Phase "permissions-entra"
            Write-Log "WARN" "Erreur permission Entra"
        }
    }
}

# =========================
# COPIE D'UN ITEM
# =========================
function Copy-One-REAL {
    param(
        [string]$Rel,
        [string]$SrcDriveId, [string]$SrcToken,
        [string]$DstDriveId, [string]$DstToken
    )

    $originalRel = Normalize-RelPath (Strip-LibraryPrefix (Normalize-RelPath $Rel))
    if ([string]::IsNullOrWhiteSpace($originalRel)) {
        return @{ Ok = $false; ItemId = $null; TranslatedRel = "" }
    }

    $translatedRel = Normalize-RelPath (Apply-UrlTranslations -OriginalUrl $originalRel)
    Write-Log "INFO" ("Copie: {0}" -f $originalRel)

    try {
        $dst = Item-Exists -DriveId $DstDriveId -Rel $translatedRel -AccessToken $DstToken
        if ($dst.Exists -and -not $ForceOverwrite) {
            return @{ Ok = $true; ItemId = $dst.Id; TranslatedRel = $translatedRel }
        }

        $usedSrcDriveId = $SrcDriveId
        $src = Item-Exists -DriveId $usedSrcDriveId -Rel $originalRel -AccessToken $SrcToken

        if (-not $src.Exists -and $script:SrcAllDrives -and $script:SrcAllDrives.Count -gt 0) {
            foreach ($d in $script:SrcAllDrives) {
                if ($d.id -eq $usedSrcDriveId) { continue }
                $probe = Item-Exists -DriveId $d.id -Rel $originalRel -AccessToken $SrcToken
                if ($probe.Exists) {
                    $usedSrcDriveId = $d.id
                    $src = $probe
                    Write-Log "INFO" "Source trouvée dans une autre bibliothèque"
                    break
                }
            }
        }

        if (-not $src.Exists) {
            Append-Err -Rel $originalRel -ErrMsg "Source introuvable" -Phase "lookup"
            return @{ Ok = $false; ItemId = $null; TranslatedRel = $translatedRel }
        }

        if ($src.IsFolder) {
            $folderId = Ensure-FolderPath -DriveId $DstDriveId -FolderRel $translatedRel -AccessToken $DstToken
            Write-Log "SUCCESS" ("Dossier créé: {0}" -f $translatedRel)
            return @{ Ok = $true; ItemId = $folderId; TranslatedRel = $translatedRel }
        }

        if ($CopierVersions) {
            $copyResult = Copy-FileWithVersions -Rel $originalRel -SrcDriveId $usedSrcDriveId -SrcToken $SrcToken -DstDriveId $DstDriveId -DstToken $DstToken -SrcConnexion $connexion_source -DstConnexion $connexion_cible
            if ($copyResult.Ok) {
                Write-Log "SUCCESS" ("Copié avec versions: {0}" -f $translatedRel)
                return $copyResult
            }
        }

        $srcTokRef = [ref]$SrcToken
        $bytes = Download-FileContent -DriveId $usedSrcDriveId -RelativePath $originalRel -Connexion $connexion_source -TokenRef $srcTokRef

        $dstTokRef = [ref]$DstToken
        $uploaded = if ($bytes.Length -le 4MB) {
            Upload-SmallFile -DriveId $DstDriveId -RelativePath $translatedRel -Bytes $bytes -Connexion $connexion_cible -TokenRef $dstTokRef
        } else {
            Upload-LargeFile -DriveId $DstDriveId -RelativePath $translatedRel -Bytes $bytes -Connexion $connexion_cible -TokenRef $dstTokRef -ChunkBytes ([int]$ChunkSize)
        }

        $itemId = if ($uploaded -and $uploaded.id) { $uploaded.id } else { $null }
        if (-not $itemId) {
            $check = Item-Exists -DriveId $DstDriveId -Rel $translatedRel -AccessToken $dstTokRef.Value
            if ($check.Exists) { $itemId = $check.Id }
        }

        Write-Log "SUCCESS" ("Copié: {0}" -f $translatedRel)
        return @{ Ok = $true; ItemId = $itemId; TranslatedRel = $translatedRel }

    } catch {
        $detail = Get-ErrorDetail $_
        Append-Err -Rel $originalRel -ErrMsg $detail -Phase "copy"
        Write-Log "ERROR" ("Erreur copie: {0}" -f $detail)
        return @{ Ok = $false; ItemId = $null; TranslatedRel = $translatedRel }
    }
}

# =========================
# MAIN
# =========================
try {
    Test-ConnexionSchema -Conn $connexion_source -Name "source"
    Test-ConnexionSchema -Conn $connexion_cible -Name "cible"

    Write-Log "INFO" "========================================="
    Write-Log "SUCCESS" "=== COPIE SHAREPOINT + PERMISSIONS + VERSIONS ==="
    Write-Log "INFO" "Version: 9.3"
    Write-Log "INFO" ("Log:    {0}" -f $script:LogPath)
    Write-Log "INFO" ("Errors: {0}" -f $script:ErrCsvPath)

    Load-UrlTranslations -TranslationFile $translation_csv
    Load-UserMapping -MappingFile $mapping_csv

    Write-Log "INFO" "Authentification..."
    $srcToken = Get-AccessToken -TenantId $connexion_source.tenant_id -ClientId $connexion_source.client_id -ClientSecret $connexion_source.client_secret
    $dstToken = Get-AccessToken -TenantId $connexion_cible.tenant_id -ClientId $connexion_cible.client_id -ClientSecret $connexion_cible.client_secret

    Write-Log "INFO" "Résolution sites/drives..."
    try {
        $srcCtx = Get-SiteAndDrive -SiteUrl $site_url_source -AccessToken $srcToken
    } catch {
        throw "Résolution site source en échec ($site_url_source) : $($_.Exception.Message)"
    }
    try {
        $dstCtx = Get-SiteAndDrive -SiteUrl $site_url_cible -AccessToken $dstToken
    } catch {
        throw "Résolution site cible en échec ($site_url_cible) : $($_.Exception.Message)"
    }

    if (-not $srcCtx.Drive) { throw "Drive source introuvable." }
    if (-not $dstCtx.Drive) { throw "Drive destination introuvable." }

    Write-Log "SUCCESS" "Sites/drives résolus"

    $script:SrcAllDrives = Get-AllDrivesFromSite -SiteId $srcCtx.Site.id -AccessToken $srcToken

    $rows = Import-Csv -LiteralPath $chemin_csv -Encoding UTF8 | Sort-Object Path
    Write-Log "SUCCESS" ("Lignes CSV: {0}" -f $rows.Count)

    if (-not (Test-Path $mapping_csv)) {
        Generate-UserMappingTemplate -Rows $rows -MappingFile $mapping_csv
        Write-Log "WARN" "Template mapping créé. Complète DestinationDisplayName puis relance."
    }

    $grouped = $rows | Group-Object -Property Path
    Write-Log "SUCCESS" ("Chemins uniques: {0}" -f $grouped.Count)

    foreach ($group in $grouped) {
        $rel = $group.Name
        $permRows = $group.Group

        $srcToken = Get-AccessToken -TenantId $connexion_source.tenant_id -ClientId $connexion_source.client_id -ClientSecret $connexion_source.client_secret
        $dstToken = Get-AccessToken -TenantId $connexion_cible.tenant_id -ClientId $connexion_cible.client_id -ClientSecret $connexion_cible.client_secret

        Write-Log "INFO" ("Traitement: {0}" -f $rel)

        $copy = Copy-One-REAL -Rel $rel -SrcDriveId $srcCtx.Drive.id -SrcToken $srcToken -DstDriveId $dstCtx.Drive.id -DstToken $dstToken

        if ($copy.Ok) {
            $dstTokRef = [ref]$dstToken
            Apply-Item-Permissions -TranslatedRel $copy.TranslatedRel -DstDriveId $dstCtx.Drive.id -DstSiteUrl $site_url_cible -DstTokenRef $dstTokRef -PermRows $permRows -KnownItemId $copy.ItemId
        } else {
            Write-Log "WARN" ("Droits ignorés (copie échouée): {0}" -f $rel)
        }
    }

    Write-Log "SUCCESS" "=== TERMINÉ ==="
}
catch {
    Write-Log "ERROR" ("Erreur fatale: {0}" -f $_.Exception.Message)
    exit 1
}
finally {
    if ($script:ErrRows.Count -gt 0) {
        try {
            $script:ErrRows | Export-Csv -LiteralPath $script:ErrCsvPath -NoTypeInformation -Encoding UTF8
            Write-Log "WARN" ("Erreurs exportées: {0}" -f $script:ErrCsvPath)
        } catch {}
    }
    if ($script:HttpClient) { $script:HttpClient.Dispose() }
}