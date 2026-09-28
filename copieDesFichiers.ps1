
# Script PowerShell pour copier des fichiers entre deux tenants SharePoint
# Version: 5.5 - Fix TokenRef, Download/Upload, Ensure-FolderPath, fallback bibliothÃ¨ques

# =========================
# PARAMÃˆTRES EN DUR
# =========================
#$tenantId = "85a5a352-25d6-4894-a97d-221cd1712dd2"
#$clientId = "684f670e-6931-4120-b053-f8c9538744f6"
#$clientSecret = "eRI8Q~qUTY1~ldJ6CIGyS_vF-SFnQrVZih~l6cuv"

# Configuration des connexions SharePoint (REMPLACER PAR VOS VRAIES VALEURS)
$connexion_source = @{
        tenant_Id = "85a5a352-25d6-4894-a97d-221cd1712dd2"
        client_Id = "684f670e-6931-4120-b053-f8c9538744f6"
        client_Secret = "tropsecret"
}

$connexion_cible = @{
    tenant_id     = "2eea08b8-1972-447b-ad43-d044d042500a"
    client_id     = "821109a4-6e0e-48e8-b477-8f9b70aa32a4"
    client_secret = "topsecret"
   # client_secret ="66A8Q~t2-Q_vZp54cDnf8TH~3kl0vp1UyHMZ8bVu"
}

# URLs des sites SharePoint
$site_url_source = "https://ipdlab.sharepoint.com/sites/SERVICES_GENERAUX/"
$site_url_cible  = "https://infoprodigital365.sharepoint.com/sites/fethygate/"
# Fichiers CSV
$chemin_csv      = "E:\SharePoint_Permissions_Export_src9999.csv"       # colonnes: Path,ItemType,GrantedTo,grantedToID,TargetType,Role
$translation_csv = "E:\UKREiiF_FinalCopy_path.csv"    # 2 colonnes: Source,Destination
$mapping_csv     = "E:\UserMapping.csv"                # colonnes: SourceId,SourceDisplayName,TargetType,DestinationId

# Options de copie
$ForceOverwrite = $true
$ChunkSize      = 10MB

# =========================
# VARIABLES GLOBALES
# =========================
$script:RunId      = (Get-Date -Format "yyyy-MM-dd_HHmmss")
$script:LogPath    = Join-Path -Path $PSScriptRoot -ChildPath ("\log\SharePoint_Copy_{0}.log" -f $script:RunId)
$script:ErrCsvPath = Join-Path -Path $PSScriptRoot -ChildPath ("\log\SharePoint_Errors_{0}.csv" -f $script:RunId)
$script:ErrRows    = New-Object System.Collections.Generic.List[object]
$global:UrlTranslations = @{}
$global:UserMapping    = @{}   # SourceId -> DestinationDisplayName
$script:ResolvedIdCache = @{}  # DisplayName -> ObjectId (cache pour éviter les appels répétés)
$script:TokenCache = @{}

# HttpClient
Add-Type -AssemblyName System.Net.Http
$script:HttpClient = [System.Net.Http.HttpClient]::new()
$script:HttpClient.Timeout = [TimeSpan]::FromMinutes(30)

# =========================
# LOGGING
# =========================
function Write-Log {
    param(
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG')][string]$Level = 'INFO',
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[{0}] [{1}] {2}" -f $timestamp, $Level, $Message
    switch ($Level) {
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
        "ERROR"   { Write-Host $logEntry -ForegroundColor Red }
        "WARN"    { Write-Host $logEntry -ForegroundColor Yellow }
        "INFO"    { Write-Host $logEntry -ForegroundColor Cyan }
        "DEBUG"   { Write-Host $logEntry -ForegroundColor Gray }
        default   { Write-Host $logEntry -ForegroundColor White }
    }
    Add-Content -Path $script:LogPath -Value $logEntry -Encoding UTF8
}
function Append-Err {
    param([string]$Rel, [string]$ErrMsg, [string]$Phase = '')
    $script:ErrRows.Add([pscustomobject]@{
        horodatage     = Get-Date
        chemin_relatif = $Rel
        phase          = $Phase
        erreur         = $ErrMsg
    }) | Out-Null
}

# =========================
# TRADUCTIONS D'URL
# =========================
function Load-UrlTranslations {
    param([string]$TranslationFile)
    try {
        Write-Log "INFO" ("Chargement traductions: {0}" -f $TranslationFile)
        if (-not (Test-Path $TranslationFile)) { Write-Log "WARN" "Fichier de traduction introuvable"; return }
        $csv = Import-Csv $TranslationFile -Header "Source","Destination"
        $n = 0
        foreach ($r in $csv) {
            $src = $r.Source.Trimt().Trim('"')
            $dst = $r.Destination.Trim().Trim('"')
			Write-Log "SUCCESS" ("source:{0} destination: {1}" -f $src, $dst)
            if ($src -and $dst) { $global:UrlTranslations[$src] = $dst; $n++ }
        }
        Write-Log "SUCCESS" ("Traductions chargées: {0}" -f $n)
    } catch { Write-Log "ERROR" ("Erreur chargement traductions: {0}" -f $_.Exception.Message) }
}
function Apply-UrlTranslations {
    param([string]$OriginalUrl)
    if ([string]::IsNullOrWhiteSpace($OriginalUrl)) { return $OriginalUrl }
    if ($global:UrlTranslations.Count -eq 0) { return $OriginalUrl }
    $t = $OriginalUrl
    foreach ($kv in $global:UrlTranslations.GetEnumerator()) {
        if ($t.Contains($kv.Key)) { $t = $t.Replace($kv.Key, $kv.Value) }
    }
    return $t
}

# =========================
# OUTILS
# =========================
function UrlEncode-Segment([string]$seg){
    if ([string]::IsNullOrWhiteSpace($seg)) { return "" }
    $normalized = $seg.Normalize([System.Text.NormalizationForm]::FormC)
    [System.Uri]::EscapeDataString($normalized)
}
function Encode-RelForGraph([string]$rel){
    $safeRel = Normalize-RelPath $rel
    $segs = ($safeRel -replace '\\','/') -split '/' | Where-Object { $_ -ne '' }
    $enc  = $segs | ForEach-Object { UrlEncode-Segment $_ }
    return [string]::Join('/', $enc)
}
function Normalize-RelPath([string]$rel){
    if (-not $rel) { return "" }
    $r = $rel.Normalize([System.Text.NormalizationForm]::FormC).Trim()
    $r = $r.TrimStart('\','/').TrimEnd(' ','/','\')
    return $r
}
function Build-AuthHeader([string]$token){ @{ "Authorization" = "Bearer $token" } }
function Get-HostAndPathFromSiteUrl([string]$siteUrl){ $u=[Uri]$siteUrl; [pscustomobject]@{ host=$u.Host; path=$u.AbsolutePath } }
function Get-RetryDelayMs($attempt,$retryAfter){
    if($retryAfter){$s=0; if([int]::TryParse($retryAfter,[ref]$s)){return [int]($s*1000)}}
    $b=[math]::Pow(2,[math]::Max(0,$attempt-1))*500; $j=Get-Random -Min 0 -Max 300
    [int]([math]::Min(32000,$b+$j))
}
function Is-TokenErrorText([string]$txt){
    ($txt -match '401' -or $txt -match 'invalid[_-]?token' -or $txt -match 'expired' -or $txt -match 'invalid[_-]?grant' -or $txt -match 'AADSTS')
}

# =========================
# TOKENS
# =========================
function Get-AccessToken {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret
    )
    $cacheKey = "$TenantId|$ClientId"
    $now = Get-Date
    $skew = [TimeSpan]::FromMinutes(5)
    if ($script:TokenCache.ContainsKey($cacheKey)) {
        $entry = $script:TokenCache[$cacheKey]
        if ($entry.expires -gt $now.Add($skew)) { return $entry.token }
        Write-Log "WARN" "Token presque expiré pour $cacheKey, régénération"
    }
    $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{ client_id=$ClientId; client_secret=$ClientSecret; grant_type="client_credentials"; scope="https://graph.microsoft.com/.default" }
    try {
        $resp = Invoke-RestMethod -Method POST -Uri $tokenUri -Body $body -ContentType "application/x-www-form-urlencoded"
        $token = $resp.access_token
        $expirationTime = (Get-Date).AddSeconds([int]$resp.expires_in)
        $script:TokenCache[$cacheKey] = @{ token=$token; expires=$expirationTime }
        Write-Log "SUCCESS" "Nouveau token pour $cacheKey. Expire à $expirationTime"
        return $token
    } catch {
        $msg = $_.Exception.Message; if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $msg = $_.ErrorDetails.Message }
        Write-Log "ERROR" "Échec token $cacheKey : $msg"; throw
    }
}

# =========================
# APPELS GRAPH AVEC RETRY
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
    for ($attempt=1; $attempt -le $MaxRetries; $attempt++) {
        $TokenRef.Value = Get-AccessToken -TenantId $Connexion.tenant_id -ClientId $Connexion.client_id -ClientSecret $Connexion.client_secret
        try {
            $headers = Build-AuthHeader $TokenRef.Value
            $payload = $Body
            if ($payload -ne $null -and $ContentType -eq "application/json" -and ($payload -isnot [string])) { $payload = ($payload | ConvertTo-Json -Depth 12) }
            return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $payload -ContentType $ContentType
        } catch {
            $resp = $_.Exception.Response
            $status = $null; $retryAfter = $null
            if ($resp) { try { $status = [int]$resp.StatusCode } catch {}; if ($resp.Headers["Retry-After"]) { $retryAfter = $resp.Headers["Retry-After"] } }
            $msg = $_.Exception.Message
            $shouldRetry = ($status -eq 429) -or ($status -ge 500 -and $status -lt 600) -or ($status -eq 408) -or (Is-TokenErrorText $msg)
            if ($shouldRetry -and $attempt -lt $MaxRetries) { $delay = Get-RetryDelayMs $attempt $retryAfter; Write-Log "WARN" "Retry $attempt ($status) sur $Uri. Attente ${delay}ms"; Start-Sleep -Milliseconds $delay; continue }
            throw
        }
    }
}

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
    for ($attempt=1; $attempt -le $MaxRetries; $attempt++) {
        $TokenRef.Value = Get-AccessToken -TenantId $Connexion.tenant_id -ClientId $Connexion.client_id -ClientSecret $Connexion.client_secret
        $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $uri)
        $req.Headers.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $TokenRef.Value)
        try {
            $resp = $script:HttpClient.SendAsync($req).Result
            if ($resp.IsSuccessStatusCode) { return $resp.Content.ReadAsByteArrayAsync().Result }
            $status = [int]$resp.StatusCode
            $retryAfter = $null; if ($resp.Headers.RetryAfter -and $resp.Headers.RetryAfter.Delta) { $retryAfter = $resp.Headers.RetryAfter.Delta.TotalSeconds }
            if (($status -eq 401 -or $status -eq 403 -or $status -eq 429 -or ($status -ge 500 -and $status -lt 600)) -and $attempt -lt $MaxRetries) {
                $delay = Get-RetryDelayMs $attempt $retryAfter; Write-Log "WARN" "Download retry $attempt ($status) $RelativePath. Attente ${delay}ms"; Start-Sleep -Milliseconds $delay; continue
            }
            throw "HTTP $status"
        } catch {
            if ($attempt -lt $MaxRetries) { $delay = Get-RetryDelayMs $attempt $null; Start-Sleep -Milliseconds $delay; continue }
            throw
        }
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
    if ($rel) {
        $parent = [System.IO.Path]::GetDirectoryName($rel) -replace '\\','/'
        if ($parent -and $parent -ne ".") {
            $TokenRef.Value = Get-AccessToken -TenantId $Connexion.tenant_id -ClientId $Connexion.client_id -ClientSecret $Connexion.client_secret
            Ensure-FolderPath -DriveId $DriveId -FolderRel $parent -AccessToken $TokenRef.Value
        }
    }
    $encRel = Encode-RelForGraph $rel
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$encRel`:/content"
    $TokenRef.Value = Get-AccessToken -TenantId $Connexion.tenant_id -ClientId $Connexion.client_id -ClientSecret $Connexion.client_secret
    $headers = Build-AuthHeader $TokenRef.Value
    return Invoke-RestMethod -Method PUT -Uri $uri -Headers $headers -Body $Bytes -ContentType "application/octet-stream"
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
    $encRel = Encode-RelForGraph $RelativePath
    $createUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$encRel`:/createUploadSession"
    $session = Invoke-GraphJson -Method POST -Uri $createUri -Connexion $Connexion -TokenRef $TokenRef -Body @{ item = @{ "@microsoft.graph.conflictBehavior" = "replace" } }
    $uploadUrl = $session.uploadUrl

    $total = $Bytes.Length
    $pos = 0
    $attempt = 0
    while ($pos -lt $total) {
        $attempt++
        $chunk = [Math]::Min($ChunkBytes, $total - $pos)
        $from = $pos
        $to   = $pos + $chunk - 1

        $content = [System.Net.Http.ByteArrayContent]::new($Bytes, $from, $chunk)
        $content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::new("application/octet-stream")
        $null = $content.Headers.TryAddWithoutValidation("Content-Range", ("bytes {0}-{1}/{2}" -f $from, $to, $total))

        $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Put, $uploadUrl)
        $req.Content = $content

        $resp = $script:HttpClient.SendAsync($req).Result
        if ($resp.IsSuccessStatusCode) {
            $pos += $chunk
        } else {
            $status = [int]$resp.StatusCode
            $retryAfter = $null
            if ($resp.Headers.RetryAfter -and $resp.Headers.RetryAfter.Delta) { $retryAfter = $resp.Headers.RetryAfter.Delta.TotalSeconds }
            if ($status -eq 429 -or ($status -ge 500 -and $status -lt 600)) {
                $delay = Get-RetryDelayMs $attempt $retryAfter
                Write-Log "WARN" "Chunk retry ($status) $RelativePath bytes $from-$to. Attente ${delay}ms"
                Start-Sleep -Milliseconds $delay
                continue
            } else {
                $txt = $resp.Content.ReadAsStringAsync().Result
                throw "Upload échec HTTP $status : $txt"
            }
        }
    }
    return $true
}

# =========================
# FONCTIONS SP
# =========================
function Get-SiteAndDrive {
    param([string]$SiteUrl, [string]$AccessToken)
    $hp = Get-HostAndPathFromSiteUrl $SiteUrl
    $siteUri = "https://graph.microsoft.com/v1.0/sites/$($hp.host):$($hp.path)?`$select=id,webUrl,displayName"
    $headers = Build-AuthHeader $AccessToken
    $site = Invoke-RestMethod -Method GET -Uri $siteUri -Headers $headers

    $drivesUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives?`$select=id,name,driveType"
    $drives = Invoke-RestMethod -Method GET -Uri $drivesUri -Headers $headers

    $drive = $drives.value | Where-Object { $_.driveType -eq 'documentLibrary' -and $_.name -eq 'BEN_RestoredFolders' } | Select-Object -First 1
    if (-not $drive) { $drive = $drives.value | Where-Object { $_.driveType -eq 'documentLibrary' } | Select-Object -First 1 }

    [pscustomobject]@{ Site=$site; Drive=$drive }
}
function Get-AllDrivesFromSite {
    param([string]$SiteId, [string]$AccessToken)
    $headers = @{ Authorization = "Bearer $AccessToken" }
    $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives?`$select=id,name,driveType"
    $r = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
    $r.value | Where-Object { $_.driveType -eq 'documentLibrary' }
}
function Item-Exists {
    param([string]$DriveId, [string]$Rel, [string]$AccessToken)
    $encRel = Encode-RelForGraph $Rel
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$encRel"
    $headers = Build-AuthHeader $AccessToken
    try {
        $it = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
        $isFolder = ($null -ne $it.folder)
        [pscustomobject]@{ Exists=$true; IsFolder=$isFolder; Id=$it.id; Size=$it.size }
    } catch {
        if ($_.Exception.Response -and [int]$_.Exception.Response.StatusCode -eq 404) {
            [pscustomobject]@{ Exists=$false; IsFolder=$false; Id=$null; Size=0 }
        } else { throw }
    }
}
function Ensure-FolderPath {
    param([string]$DriveId, [string]$FolderRel, [string]$AccessToken)
    $rel  = Normalize-RelPath $FolderRel
    $segs = $rel -split '[\\/]' | Where-Object { $_ -ne '' }
    if ($segs.Count -eq 0) { return }

    $headers = Build-AuthHeader $AccessToken
    $currentPath = ""

    foreach ($s in $segs) {
        if ([string]::IsNullOrEmpty($currentPath)) { $currentPath = $s } else { $currentPath = "$currentPath/$s" }

        $enc = Encode-RelForGraph $currentPath
        $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$enc"
        $exists = $true
        try { [void](Invoke-RestMethod -Method GET -Uri $uri -Headers $headers) } catch {
            if ($_.Exception.Response -and [int]$_.Exception.Response.StatusCode -eq 404) { $exists = $false } else { throw }
        }

        if (-not $exists) {
            $parentPath = $currentPath.Substring(0, $currentPath.Length - $s.Length).TrimEnd('/')
            if ([string]::IsNullOrEmpty($parentPath)) {
                $createUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root/children"
            } else {
                $parentEnc = Encode-RelForGraph $parentPath
                $createUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$parentEnc`:/children"
            }
            $body = @{ name = $s.Normalize([System.Text.NormalizationForm]::FormC); folder = @{}; "@microsoft.graph.conflictBehavior" = "replace" } | ConvertTo-Json
            try {
                [void](Invoke-RestMethod -Method POST -Uri $createUri -Headers $headers -Body $body -ContentType "application/json" -ErrorAction Stop)
            } catch {
                $detail = $_.Exception.Message
                if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $detail = $_.ErrorDetails.Message }
                throw "Création dossier impossible '$currentPath' : $detail"
            }
        }
    }
}

# =========================
# MAPPAGE UTILISATEURS
# =========================
function Generate-UserMappingTemplate {
    param([object[]]$Rows, [string]$MappingFile)
    $unique = $Rows | Where-Object { $_.grantedToID -and $_.grantedToID -ne 'Unknown' } |
              Sort-Object grantedToID -Unique |
              Select-Object grantedToID, GrantedTo, TargetType
    $template = $unique | ForEach-Object {
        [pscustomobject]@{
            SourceId               = $_.grantedToID
            SourceDisplayName      = $_.GrantedTo
            TargetType             = $_.TargetType
            DestinationDisplayName = ""   # <-- Nom de l'utilisateur/groupe dans le tenant destination
        }
    }
    $template | Export-Csv -Path $MappingFile -NoTypeInformation -Encoding UTF8
    Write-Log "INFO" ("Modèle de mappage généré: {0} ({1} entrées à compléter)" -f $MappingFile, $template.Count)
    Write-Log "WARN" "Remplissez la colonne DestinationId dans $MappingFile puis relancez le script."
}

function Load-UserMapping {
    param([string]$MappingFile)
    if (-not (Test-Path $MappingFile)) {
        Write-Log "WARN" ("Fichier de mappage utilisateurs introuvable: {0} - les droits ne seront pas appliqués" -f $MappingFile)
        return
    }
    $csv = Import-Csv $MappingFile -Encoding UTF8
    $n = 0
    foreach ($row in $csv) {
        $srcId   = $row.SourceId.Trim()
        $dstName = $row.DestinationDisplayName.Trim()
        if ($srcId -and $dstName) { $global:UserMapping[$srcId] = @{ DisplayName = $dstName; TargetType = $row.TargetType.Trim() }; $n++ }
    }
    Write-Log "SUCCESS" ("Mappages utilisateurs chargés: {0}" -f $n)
}

function Resolve-DestinationObjectId {
    param(
        [string]$DisplayName,
        [string]$TargetType,
        [ref]$DstTokenRef
    )
    $cacheKey = "$TargetType|$DisplayName"
    if ($script:ResolvedIdCache.ContainsKey($cacheKey)) { return $script:ResolvedIdCache[$cacheKey] }

    # Utiliser $search avec ConsistencyLevel:eventual (ne nécessite que User.Read.All en lecture)
    $encoded = [System.Uri]::EscapeDataString($DisplayName)
    if ($TargetType -eq 'Group') {
        $uri = "https://graph.microsoft.com/v1.0/groups?`$search=`"displayName:$encoded`"&`$select=id,displayName"
    } else {
        $uri = "https://graph.microsoft.com/v1.0/users?`$search=`"displayName:$encoded`"&`$select=id,displayName,userPrincipalName"
    }

    try {
        $token = Get-AccessToken -TenantId $connexion_cible.tenant_id -ClientId $connexion_cible.client_id -ClientSecret $connexion_cible.client_secret
        $headers = Build-AuthHeader $token
        $headers['ConsistencyLevel'] = 'eventual'
        $result = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
        # Chercher correspondance exacte parmi les résultats
        $obj = $result.value | Where-Object { $_.displayName -eq $DisplayName } | Select-Object -First 1
        if (-not $obj) { $obj = $result.value | Select-Object -First 1 }
        if ($obj) {
            $script:ResolvedIdCache[$cacheKey] = $obj.id
            Write-Log "SUCCESS" ("Résolution: '{0}' ({1}) -> {2}" -f $DisplayName, $TargetType, $obj.id)
            return $obj.id
        } else {
            Write-Log "WARN" ("Aucun {0} trouvé avec le nom '{1}' dans le tenant destination" -f $TargetType, $DisplayName)
            $script:ResolvedIdCache[$cacheKey] = $null
            return $null
        }
    } catch {
        $statusCode = $null
        if ($_.Exception.Response) { try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {} }
        if ($statusCode -eq 403) {
            Write-Log "WARN" ("[403] Permissions insuffisantes pour résoudre '{0}'. Vérifiez que l'app destination a User.Read.All ou GroupMember.Read.All." -f $DisplayName)
        } else {
            Write-Log "WARN" ("Erreur résolution '{0}': {1}" -f $DisplayName, $_.Exception.Message)
        }
        $script:ResolvedIdCache[$cacheKey] = $null
        return $null
    }
}

function Apply-Item-Permissions {
    param(
        [string]$TranslatedRel,
        [string]$DstDriveId,
        [ref]$DstTokenRef,
        [object[]]$PermRows
    )
    if (-not $PermRows -or $PermRows.Count -eq 0) { return }
    if ($global:UserMapping.Count -eq 0) { Write-Log "WARN" "Aucun mappage chargé - droits ignorés pour: $TranslatedRel"; return }

    # Récupérer l'ID de l'item destination (retry si 404 = race condition après création dossier)
    $encRel = Encode-RelForGraph $TranslatedRel
    $itemUri = "https://graph.microsoft.com/v1.0/drives/$DstDriveId/root:/$encRel"
    $itemId = $null
    for ($ri = 1; $ri -le 4; $ri++) {
        try {
            $item = Invoke-GraphJson -Method GET -Uri $itemUri -Connexion $connexion_cible -TokenRef $DstTokenRef
            $itemId = $item.id
            break
        } catch {
            $st = $null; if ($_.Exception.Response) { try { $st = [int]$_.Exception.Response.StatusCode } catch {} }
            if ($st -eq 404 -and $ri -lt 4) {
                Write-Log "DEBUG" ("[404] Item pas encore disponible, retry {0}/3 dans 2s: {1}" -f $ri, $TranslatedRel)
                Start-Sleep -Seconds 2
            } else {
                Write-Log "WARN" ("Impossible de récupérer l'item destination pour les droits: {0} - {1}" -f $TranslatedRel, $_.Exception.Message)
                return
            }
        }
    }
    if (-not $itemId) { return }

    foreach ($perm in $PermRows) {
        $srcId = $perm.grantedToID
        if (-not $global:UserMapping.ContainsKey($srcId)) {
            Write-Log "WARN" ("Pas de mappage pour: {0} ({1}) - permission ignorée sur {2}" -f $perm.GrantedTo, $srcId, $TranslatedRel)
            continue
        }
        $entry = $global:UserMapping[$srcId]
        $dstId = Resolve-DestinationObjectId -DisplayName $entry.DisplayName -TargetType $entry.TargetType -DstTokenRef $DstTokenRef
        if (-not $dstId) {
            Write-Log "WARN" ("Impossible de résoudre '{0}' dans le tenant destination - permission ignorée" -f $entry.DisplayName)
            continue
        }

        # L'endpoint /invite ne supporte que read et write - owner doit être géré via le groupe site
        $role = switch ($perm.Role.ToLower()) {
            "owner"  { "write" }   # /invite ne supporte pas owner -> write
            "write"  { "write" }
            "read"   { "read" }
            default  { "read" }
        }

        $inviteUri = "https://graph.microsoft.com/v1.0/drives/$DstDriveId/items/$itemId/invite"
        $body = @{
            requireSignIn  = $true
            sendInvitation = $false
            roles          = @($role)
            recipients     = @(@{ objectId = $dstId })
        }
        try {
            Invoke-GraphJson -Method POST -Uri $inviteUri -Connexion $connexion_cible -TokenRef $DstTokenRef -Body $body
            Write-Log "SUCCESS" ("Permission appliquée: {0} [{1}] -> {2}" -f $perm.GrantedTo, $role, $TranslatedRel)
        } catch {
            Write-Log "WARN" ("Erreur permission {0} [{1}] sur {2}: {3}" -f $perm.GrantedTo, $role, $TranslatedRel, $_.Exception.Message)
            Append-Err -Rel $TranslatedRel -ErrMsg ("Permission {0} [{1}]: {2}" -f $perm.GrantedTo, $role, $_.Exception.Message) -Phase "permissions"
        }
    }
}

# =========================
# COPIE
# =========================
function Copy-One-REAL {
    param(
        [string]$Rel,
        [string]$SrcDriveId, [string]$SrcToken,
        [string]$DstDriveId, [string]$DstToken
    )
    $originalRel = Normalize-RelPath $Rel
    if ($originalRel -eq "") { return $false }
    Write-Log "INFO" ("=== COPIE REELLE: {0} ===" -f $originalRel)

    $translatedRel = Apply-UrlTranslations -OriginalUrl $originalRel
    if ($translatedRel -ne $originalRel) { Write-Log "SUCCESS" ("URL traduite: {0} -> {1}" -f $originalRel, $translatedRel) }

    $maxTokenRetries = 1
    for ($attempt = 0; $attempt -le $maxTokenRetries; $attempt++) {
        try {
            $dst = Item-Exists -DriveId $DstDriveId -Rel $translatedRel -AccessToken $DstToken
            if ($dst.Exists -and -not $ForceOverwrite) { Write-Log "INFO" ("Fichier existant, ignore: {0}" -f $translatedRel); return }

            $src = Item-Exists -DriveId $SrcDriveId -Rel $originalRel -AccessToken $SrcToken
            if (-not $src.Exists -and $script:SrcAllDrives -and $script:SrcAllDrives.Count -gt 0) {
                foreach ($d in $script:SrcAllDrives) {
                    if ($d.id -eq $SrcDriveId) { continue }
                    $probe = Item-Exists -DriveId $d.id -Rel $originalRel -AccessToken $SrcToken
                    if ($probe.Exists) { Write-Log "SUCCESS" ("Chemin trouvé dans la bibliothèque source '{0}'" -f $d.name); $SrcDriveId = $d.id; $src = $probe; break }
                }
            }
            if (-not $src.Exists) { Append-Err -Rel $originalRel -ErrMsg "Fichier/Dossier source introuvable (toutes bibliothèques testées)" -Phase "lookup"; Write-Log "ERROR" ("Fichier source introuvable: {0}" -f $originalRel); return $false }

            if ($src.IsFolder) { Ensure-FolderPath -DriveId $DstDriveId -FolderRel $translatedRel -AccessToken $DstToken; Write-Log "SUCCESS" ("Dossier cree: {0}" -f $translatedRel); return $true }

            $srcTokRef = [ref]$SrcToken
            $bytes = Download-FileContent -DriveId $SrcDriveId -RelativePath $originalRel -Connexion $connexion_source -TokenRef $srcTokRef

            $dstTokRef = [ref]$DstToken
            if ($bytes.Length -le 4MB) {
                [void](Upload-SmallFile -DriveId $DstDriveId -RelativePath $translatedRel -Bytes $bytes -Connexion $connexion_cible -TokenRef $dstTokRef)
            } else {
                [void](Upload-LargeFile -DriveId $DstDriveId -RelativePath $translatedRel -Bytes $bytes -Connexion $connexion_cible -TokenRef $dstTokRef -ChunkBytes ([int]$ChunkSize))
            }

            Write-Log "SUCCESS" ("=== FICHIER COPIE AVEC SUCCES: {0} ===" -f $translatedRel)
            return $true
        } catch {
            $msg = $_.Exception.Message
            if ($attempt -lt $maxTokenRetries -and (Is-TokenErrorText $msg)) {
                Write-Log "WARN" "Token expiré/invalid → refresh et retry"
                $SrcToken = Get-AccessToken -TenantId $connexion_source.tenant_id -ClientId $connexion_source.client_id -ClientSecret $connexion_source.client_secret
                $DstToken = Get-AccessToken -TenantId $connexion_cible.tenant_id -ClientId $connexion_cible.client_id -ClientSecret $connexion_cible.client_secret
                Start-Sleep -Milliseconds 800
                continue
            }
            Append-Err -Rel $originalRel -ErrMsg $msg -Phase "copy"
            Write-Log "ERROR" ("Erreur copie: {0}" -f $msg)
            return $false
        }
    }
}

# =========================
# SCRIPT PRINCIPAL
# =========================
try {
    Write-Log "INFO" "========================================="
    Write-Log "SUCCESS" "=== COPIE SHAREPOINT - GESTION DES TOKENS ==="
    Write-Log "INFO" "Version: 5.7"
    Write-Log "INFO" "========================================="

    Load-UrlTranslations -TranslationFile $translation_csv
    if ($global:UrlTranslations.Count -gt 0) { Write-Log "INFO" ("Mappings: {0}" -f $global:UrlTranslations.Count) } else { Write-Log "WARN" "Aucune traduction chargée. Copie en chemins identiques." }

    Load-UserMapping -MappingFile $mapping_csv

    Write-Log "INFO" "=== AUTHENTIFICATION ==="
    $srcToken = Get-AccessToken -TenantId $connexion_source.tenant_id -ClientId $connexion_source.client_id -ClientSecret $connexion_source.client_secret
    $dstToken = Get-AccessToken -TenantId $connexion_cible.tenant_id -ClientId $connexion_cible.client_id -ClientSecret $connexion_cible.client_secret

    Write-Log "INFO" "=== RESOLUTION DES SITES ==="
    $srcCtx = Get-SiteAndDrive -SiteUrl $site_url_source -AccessToken $srcToken
    $dstCtx = Get-SiteAndDrive -SiteUrl $site_url_cible  -AccessToken $dstToken
    Write-Log "SUCCESS" ("Source: {0}" -f $srcCtx.Site.webUrl)
    Write-Log "SUCCESS" ("Drive source: {0} (ID: {1})" -f $srcCtx.Drive.name, $srcCtx.Drive.id)
    Write-Log "SUCCESS" ("Destination: {0}" -f $dstCtx.Site.webUrl)
    Write-Log "SUCCESS" ("Drive destination: {0} (ID: {1})" -f $dstCtx.Drive.name, $dstCtx.Drive.id)

    $script:SrcAllDrives = Get-AllDrivesFromSite -SiteId $srcCtx.Site.id -AccessToken $srcToken
    $script:DstAllDrives = Get-AllDrivesFromSite -SiteId $dstCtx.Site.id -AccessToken $dstToken
    $srcLibs = ($script:SrcAllDrives | Select-Object -ExpandProperty name) -join ", "
    $dstLibs = ($script:DstAllDrives | Select-Object -ExpandProperty name) -join ", "
    Write-Log "INFO" ("Librairies source: {0}" -f $srcLibs)
    Write-Log "INFO" ("Librairies dest:   {0}" -f $dstLibs)

    Write-Log "INFO" "=== LECTURE DU FICHIER CSV ==="
    $rows = Import-Csv -LiteralPath $chemin_csv -Encoding utf8 | Sort-Object Path
    Write-Log "SUCCESS" ("Lignes de permissions chargées: {0}" -f $rows.Count)

    # Générer le modèle de mappage si absent
    if (-not (Test-Path $mapping_csv)) {
        Generate-UserMappingTemplate -Rows $rows -MappingFile $mapping_csv
    }

    # Grouper par chemin pour ne copier chaque item qu'une seule fois
    $grouped = $rows | Group-Object -Property Path
    Write-Log "SUCCESS" ("Chemins uniques à traiter: {0}" -f $grouped.Count)

    foreach ($group in $grouped) {
        $rel      = $group.Name
        $permRows = $group.Group

        $srcToken = Get-AccessToken -TenantId $connexion_source.tenant_id -ClientId $connexion_source.client_id -ClientSecret $connexion_source.client_secret
        $dstToken = Get-AccessToken -TenantId $connexion_cible.tenant_id  -ClientId $connexion_cible.client_id  -ClientSecret $connexion_cible.client_secret

        Write-Log "INFO" ("Traitement: {0} ({1} permission(s))" -f $rel, $permRows.Count)
        $copyOk = Copy-One-REAL -Rel $rel -SrcDriveId $srcCtx.Drive.id -SrcToken $srcToken -DstDriveId $dstCtx.Drive.id -DstToken $dstToken

        # Appliquer les droits uniquement si la copie a réussi
        if ($copyOk) {
            $translatedRel = Apply-UrlTranslations -OriginalUrl (Normalize-RelPath $rel)
            $dstTokRef = [ref]$dstToken
            Apply-Item-Permissions -TranslatedRel $translatedRel -DstDriveId $dstCtx.Drive.id -DstTokenRef $dstTokRef -PermRows $permRows
        } else {
            Write-Log "WARN" ("Droits ignorés car copie échouée: {0}" -f $rel)
        }
    }

    Write-Log "SUCCESS" "=== COPIE TERMINEE ==="
} catch {
    Write-Log "ERROR" "Erreur fatale: $($_.Exception.Message)"; exit 1
} finally {
    if ($script:ErrRows.Count -gt 0) {
        try { $script:ErrRows | Export-Csv -LiteralPath $script:ErrCsvPath -Encoding UTF8 -NoTypeInformation } catch {}
        Write-Log "WARN" ("Erreurs consignées: {0} -> {1}" -f $script:ErrRows.Count, $script:ErrCsvPath)
    }
    if ($script:HttpClient) { $script:HttpClient.Dispose() }
}

