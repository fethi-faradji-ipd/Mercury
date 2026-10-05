# Script PowerShell - Copie SharePoint depuis les fichiers de decouverte (listes, bibliotheques, fichiers, droits)
# Version: 2.0 (sans base de donnees)
#
# Entrees : dossier $DiscoveryDirectory produit par DecouverteDesFichers.ps1
#           + $mapping_csv (UserMapping.csv, colonne DestinationDisplayName a completer)
# Sorties : dossier $OutputDirectory (journal + CSV de resultats)

# =========================
# PARAMETRES GRAPH
# =========================
$connexion_source = @{
    tenant_id     = "2eea08b8-1972-447b-ad43-d044d042500a"
    client_id     = "821109a4-6e0e-48e8-b477-8f9b70aa32a4"
    client_secret ="XYr8Q~jx8vDX30c97YP5M0-juBIahzni6WlDNdz8"
}

$connexion_cible = @{
    tenant_id     = "2eea08b8-1972-447b-ad43-d044d042500a"
    client_id     = "821109a4-6e0e-48e8-b477-8f9b70aa32a4"
    client_secret = "XYr8Q~jx8vDX30c97YP5M0-juBIahzni6WlDNdz8"

   # client_secret ="66A8Q~t2-Q_vZp54cDnf8TH~3kl0vp1UyHMZ8bVu"
}


$site_url_source = 'https://infoprodigital365.sharepoint.com/sites/fethygate/'
$site_url_cible = 'https://infoprodigital365.sharepoint.com/sites/fethiMercury/'

#$site_url_source = "https://ipdlab.sharepoint.com/sites/SERVICES_GENERAUX/"
#$site_url_cible  = "https://infoprodigital365.sharepoint.com/sites/WH_GPA-test/"

$translation_csv = "E:\UKREiiF_FinalCopy_path.csv"
$mapping_csv     = "E:\UserMapping.csv"
$CsvLibraryPrefixToStrip = ""

$ForceOverwrite = $true
$ChunkSize      = 8MB

# Dossiers
$baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$DiscoveryDirectory = Join-Path $baseDir 'SharePoint-Discovery'
$OutputDirectory    = Join-Path $baseDir 'SharePoint-Copy'

# Perimetre : toutes les bibliotheques et listes decouvertes.
# $true = cree sur la cible les listes/bibliotheques absentes (meme URL relative, colonnes personnalisees).
$CreateMissingLists = $true
# $true = copie les elements des listes classiques (les ID sont preserves au mieux, voir droitscustome_2.ps1).
$CopyListItems = $true
# $true = copie les valeurs des colonnes personnalisees des fichiers/dossiers et des elements.
$CopyItemFields = $true
# Nombre max d'elements factices crees pour conserver les ID d'une liste (trous d'ID dans la source).
$MaxIdPlaceholders = 5000
# Renommage eventuel d'une liste/bibliotheque : @{ "Shared Documents" = "Documents partages" } (cle = URL relative source).
# Attention : droitscustome_2.ps1 retrouve les objets par URL relative identique.
$LibraryNameMapping = @{}
# Colonnes jamais copiees (systeme / calculees par SharePoint).
$FieldSkipNames = @('ID', 'Id', 'ContentType', 'ContentTypeId', 'Attachments', 'FileLeafRef', 'FileRef', 'FileDirRef', 'LinkFilename', 'LinkFilenameNoMenu',
    'Created', 'Modified', 'Author', 'Editor', 'AppAuthor', 'AppEditor', '_UIVersionString', 'Edit', 'DocIcon', 'FolderChildCount', 'ItemChildCount',
    'LinkTitle', 'LinkTitleNoMenu', 'ComplianceAssetId', '_ColorTag', 'ParentLeafName', 'ParentVersionString')
# Groupes de colonnes natifs (deja presents dans la liste creee depuis le meme modele).
$NativeColumnGroups = @('_Hidden', 'Core Document Columns', 'Core Contact and Calendar Columns', 'Core Task and Issue Columns', 'Base Columns',
    'Document and Record Management Columns', 'Extended Columns', 'Display Template Columns', 'Enterprise Keywords Group', 'Publishing Columns')

# =========================
# ETAT GLOBAL
# =========================
$script:TokenCache = @{}
$script:UrlTranslations = @{}
$script:UserMapping = @{}
$script:UserMappingByName = @{}
$script:ResolvedIdCache = @{}
$script:WritableColumns = @{}   # id liste cible -> HashSet des colonnes modifiables
$script:ListIdMap = @{}         # id liste source -> id liste cible
$script:LogFile = $null

$script:ItemResults = New-Object System.Collections.Generic.List[object]
$script:PermResults = New-Object System.Collections.Generic.List[object]
$script:ListResults = New-Object System.Collections.Generic.List[object]
$script:ListItemResults = New-Object System.Collections.Generic.List[object]
$script:ErrorRows = New-Object System.Collections.Generic.List[object]

Add-Type -AssemblyName System.Net.Http
$script:HttpClient = [System.Net.Http.HttpClient]::new()
$script:HttpClient.Timeout = [TimeSpan]::FromMinutes(30)

# =========================
# LOG / RESULTATS
# =========================
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "[{0}][{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Write-Host $line
    if ($script:LogFile) {
        try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 } catch {}
    }
}

function Add-ErrorRow {
    param([string]$Path, [string]$Phase, [string]$Message)
    $script:ErrorRows.Add([pscustomobject]@{ Time = (Get-Date).ToString('s'); Path = $Path; Phase = $Phase; Message = $Message }) | Out-Null
}

function Add-ItemResult {
    param([string]$PathOriginal, [string]$PathTranslated, [string]$ItemType, [string]$ItemId, [string]$Status)
    $script:ItemResults.Add([pscustomobject]@{
        PathOriginal = $PathOriginal; PathTranslated = $PathTranslated; ItemType = $ItemType; TargetItemId = $ItemId; Status = $Status; Time = (Get-Date).ToString('s')
    }) | Out-Null
}

function Add-PermissionResult {
    param(
        [string]$PathOriginal, [string]$PathTranslated, [string]$ItemType,
        [string]$GrantedTo, [string]$GrantedToId, [string]$TargetType, [string]$Role,
        [string]$Status, [string]$ErrorMessage, [string]$MappedDestination, [string]$MappedDestinationId
    )
    $script:PermResults.Add([pscustomobject]@{
        PathOriginal = $PathOriginal; PathTranslated = $PathTranslated; ItemType = $ItemType
        GrantedTo = $GrantedTo; GrantedToId = $GrantedToId; TargetType = $TargetType; Role = $Role
        MappedDestination = $MappedDestination; MappedDestinationId = $MappedDestinationId
        Status = $Status; ErrorMessage = $ErrorMessage
    }) | Out-Null
}

function Add-ListResult {
    param([string]$ListKey, [string]$Status, [string]$TargetListId, [string]$ErrorMessage)
    $script:ListResults.Add([pscustomobject]@{ ListKey = $ListKey; Status = $Status; TargetListId = $TargetListId; ErrorMessage = $ErrorMessage }) | Out-Null
}

function Add-ListItemResult {
    param([string]$ListKey, [int]$SourceItemId, [string]$Status, $TargetItemId, [string]$ErrorMessage)
    $script:ListItemResults.Add([pscustomobject]@{ ListKey = $ListKey; SourceItemId = $SourceItemId; Status = $Status; TargetItemId = $TargetItemId; ErrorMessage = $ErrorMessage }) | Out-Null
}

function Save-Results {
    if (-not (Test-Path -LiteralPath $OutputDirectory)) { return }
    $script:ItemResults     | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'resultats_fichiers.csv') -NoTypeInformation -Encoding UTF8
    $script:PermResults     | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'resultats_droits.csv') -NoTypeInformation -Encoding UTF8
    $script:ListResults     | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'resultats_listes.csv') -NoTypeInformation -Encoding UTF8
    $script:ListItemResults | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'resultats_elements_listes.csv') -NoTypeInformation -Encoding UTF8
    $script:ErrorRows       | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'erreurs.csv') -NoTypeInformation -Encoding UTF8
    Write-Log ("Resultats ecrits dans {0} (fichiers={1}, droits={2}, listes={3}, elements={4}, erreurs={5})" -f $OutputDirectory, $script:ItemResults.Count, $script:PermResults.Count, $script:ListResults.Count, $script:ListItemResults.Count, $script:ErrorRows.Count)
}

# =========================
# LECTURE DES FICHIERS DE DECOUVERTE
# =========================
function Read-Jsonl {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    foreach ($line in [System.IO.File]::ReadLines($Path, [System.Text.Encoding]::UTF8)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $line | ConvertFrom-Json
    }
}

function Load-UserMapping {
    $rows = @()
    if (Test-Path -LiteralPath $mapping_csv) {
        $rows = @(Import-Csv -Path $mapping_csv -Encoding UTF8)
    } else {
        Write-Log "Mapping introuvable: $mapping_csv (aucun droit ne sera applique)" "WARN"
    }

    foreach ($row in $rows) {
        $srcId = if ($row.SourceId) { [string]$row.SourceId } else { "" }
        $srcName = if ($row.SourceDisplayName) { [string]$row.SourceDisplayName } else { "" }
        $dstName = if ($row.DestinationDisplayName) { [string]$row.DestinationDisplayName } else { "" }
        $tt = if ($row.TargetType) { [string]$row.TargetType } else { "" }

        if ([string]::IsNullOrWhiteSpace($dstName)) { continue }

        $entry = @{ DisplayName = $dstName; TargetType = $tt }
        if ($srcId) { $script:UserMapping[$srcId] = $entry }
        if ($srcName) { $script:UserMappingByName[$srcName] = $entry }
    }

    Write-Log ("Mappages charges: ID={0}, Name={1}" -f $script:UserMapping.Count, $script:UserMappingByName.Count) "SUCCESS"
}

# =========================
# GRAPH HELPERS
# =========================
function Normalize-RelPath([string]$rel) {
    if (-not $rel) { return "" }
    return (($rel.Trim() -replace '\\','/').TrimStart('/').TrimEnd('/ '))
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
    [System.Uri]::EscapeDataString($seg)
}

function Encode-RelForGraph([string]$rel) {
    $segs = ($rel -replace '\\','/') -split '/' | Where-Object { $_ -ne '' }
    return [string]::Join('/', ($segs | ForEach-Object { UrlEncode-Segment $_ }))
}

function Build-AuthHeader([string]$token) { @{ Authorization = "Bearer $token" } }

function Get-AccessTokenResource {
    param([string]$TenantId, [string]$ClientId, [string]$ClientSecret, [string]$ResourceAppIdUri)

    $cacheKey = "$TenantId|$ClientId|$ResourceAppIdUri"
    $now = Get-Date
    if ($script:TokenCache.ContainsKey($cacheKey)) {
        $e = $script:TokenCache[$cacheKey]
        if ($e.expires -gt $now.AddMinutes(5)) { return $e.token }
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
    $script:TokenCache[$cacheKey] = @{ token = $token; expires = (Get-Date).AddSeconds([int]$resp.expires_in) }
    return $token
}

function Get-AccessToken {
    param([hashtable]$Conn)
    return (Get-AccessTokenResource -TenantId $Conn.tenant_id -ClientId $Conn.client_id -ClientSecret $Conn.client_secret -ResourceAppIdUri "https://graph.microsoft.com")
}

function Invoke-GraphJson {
    param(
        [string]$Method,
        [string]$Uri,
        [hashtable]$Connexion,
        $Body = $null,
        [int]$MaxRetries = 5
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $tok = Get-AccessToken -Conn $Connexion
            $headers = Build-AuthHeader $tok

            if ($null -ne $Body) {
                return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body ($Body | ConvertTo-Json -Depth 12) -ContentType "application/json"
            }
            return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
        } catch {
            $st = $null
            try { $st = [int]$_.Exception.Response.StatusCode } catch {}
            $retry = ($st -in @(401,408,429)) -or ($st -ge 500 -and $st -lt 600)
            if ($retry -and $attempt -lt $MaxRetries) {
                Start-Sleep -Seconds ([Math]::Min(20, [Math]::Pow(2, $attempt)))
                continue
            }
            throw
        }
    }
}

function Get-HostAndPathFromSiteUrl([string]$siteUrl) {
    $u = [Uri]$siteUrl
    $p = $u.AbsolutePath
    if ($p.Length -gt 1) { $p = $p.TrimEnd('/') }
    [pscustomobject]@{ host = $u.Host; path = $p }
}

function Get-GraphAllPages {
    param([string]$Uri, [hashtable]$Connexion)

    $all = New-Object System.Collections.Generic.List[object]
    while ($Uri) {
        $r = Invoke-GraphJson -Method "GET" -Uri $Uri -Connexion $Connexion
        if (-not $r) { break }
        foreach ($v in $r.value) { $all.Add($v) | Out-Null }
        $Uri = $r.'@odata.nextLink'
    }
    return $all.ToArray()
}

# Chemin relatif de la liste dans le site, ex. "Shared Documents" ou "Lists/Taches".
function Get-ListRelativeKey {
    param([string]$WebUrl, [string]$SiteWebUrl)

    $p = [Uri]::UnescapeDataString(([Uri]$WebUrl).AbsolutePath)
    $base = [Uri]::UnescapeDataString(([Uri]$SiteWebUrl).AbsolutePath).TrimEnd('/')
    if ($base -and $p.StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase)) {
        $p = $p.Substring($base.Length)
    }
    $p = $p.Trim('/')
    $p = $p -replace '(?i)/Forms(/.*)?$', ''
    $p = $p -replace '(?i)/[^/]+\.aspx$', ''
    return $p
}

# Site + toutes ses listes/bibliotheques, indexees par URL relative (minuscules).
function Get-SiteContext {
    param([string]$SiteUrl, [hashtable]$Connexion)

    $tok = Get-AccessToken -Conn $Connexion
    $hp = Get-HostAndPathFromSiteUrl $SiteUrl

    $site = Invoke-RestMethod -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($hp.host):$($hp.path)?`$select=id,webUrl,displayName" -Headers (Build-AuthHeader $tok)
    $lists = @(Get-GraphAllPages -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists?`$select=id,name,displayName,webUrl,list" -Connexion $Connexion)

    $map = @{}
    foreach ($l in $lists) {
        $key = Get-ListRelativeKey -WebUrl $l.webUrl -SiteWebUrl $site.webUrl
        $isLib = ([string]$l.list.template -match '(?i)library$')
        $driveId = $null
        if ($isLib) {
            try {
                $d = Invoke-GraphJson -Method "GET" -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($l.id)/drive?`$select=id" -Connexion $Connexion
                $driveId = $d.id
            } catch {}
        }
        $map[$key.ToLowerInvariant()] = @{ ListId = [string]$l.id; Key = $key; IsLibrary = $isLib; DriveId = $driveId }
    }
    return [pscustomobject]@{ Site = $site; Lists = $map }
}

function Get-TargetKey([string]$SourceKey) {
    if ($LibraryNameMapping.ContainsKey($SourceKey)) { return [string]$LibraryNameMapping[$SourceKey] }
    return $SourceKey
}

# "Shared Documents/dossier/fichier.docx" -> bibliotheque + chemin relatif au drive.
function Split-LibraryPath([string]$Path) {
    $p = ([string]$Path).TrimStart('/')
    $idx = $p.IndexOf('/')
    if ($idx -lt 0) { return [pscustomobject]@{ Library = $p; Rel = "" } }
    return [pscustomobject]@{ Library = $p.Substring(0, $idx); Rel = $p.Substring($idx + 1) }
}

function Item-Exists {
    param([string]$DriveId, [string]$Rel, [string]$AccessToken)

    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$(Encode-RelForGraph $Rel)"
    try {
        $it = Invoke-RestMethod -Method GET -Uri $uri -Headers (Build-AuthHeader $AccessToken)
        return [pscustomobject]@{ Exists = $true; IsFolder = ($null -ne $it.folder); Id = $it.id }
    } catch {
        return [pscustomobject]@{ Exists = $false; IsFolder = $false; Id = $null }
    }
}

function Ensure-FolderPath {
    param([string]$DriveId, [string]$FolderRel, [string]$AccessToken)

    $rel = Normalize-RelPath $FolderRel
    $segs = $rel -split '/' | Where-Object { $_ -ne '' }
    if ($segs.Count -eq 0) { return $null }

    $parentId = "root"
    $lastId = $null

    foreach ($seg in $segs) {
        $enc = UrlEncode-Segment $seg
        $checkUri = if ($parentId -eq "root") {
            "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$enc"
        } else {
            "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$parentId`:/$enc"
        }

        $child = $null
        try {
            $child = Invoke-RestMethod -Method GET -Uri $checkUri -Headers (Build-AuthHeader $AccessToken)
        } catch {
            $createUri = if ($parentId -eq "root") {
                "https://graph.microsoft.com/v1.0/drives/$DriveId/root/children"
            } else {
                "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$parentId/children"
            }
            $body = @{ name = $seg; folder = @{}; "@microsoft.graph.conflictBehavior" = "replace" } | ConvertTo-Json -Depth 4
            $child = Invoke-RestMethod -Method POST -Uri $createUri -Headers (Build-AuthHeader $AccessToken) -Body $body -ContentType "application/json"
        }

        $parentId = $child.id
        $lastId = $child.id
    }

    return $lastId
}

function Download-FileContent {
    param([string]$DriveId, [string]$RelativePath, [hashtable]$Connexion)

    $tok = Get-AccessToken -Conn $Connexion
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$(Encode-RelForGraph $RelativePath)`:/content"

    $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $uri)
    $req.Headers.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $tok)

    $resp = $script:HttpClient.SendAsync($req).Result
    if (-not $resp.IsSuccessStatusCode) {
        throw "Download HTTP $([int]$resp.StatusCode)"
    }
    return $resp.Content.ReadAsByteArrayAsync().Result
}

function Upload-File {
    param([string]$DriveId, [string]$RelativePath, [byte[]]$Bytes, [hashtable]$Connexion)

    $rel = Normalize-RelPath $RelativePath
    $parent = ([System.IO.Path]::GetDirectoryName($rel) -replace '\\','/')
    $tok = Get-AccessToken -Conn $Connexion

    if ($parent -and $parent -ne ".") {
        Ensure-FolderPath -DriveId $DriveId -FolderRel $parent -AccessToken $tok | Out-Null
    }

    if ($Bytes.Length -le 4MB) {
        $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$(Encode-RelForGraph $rel)`:/content"
        $u = Invoke-RestMethod -Method PUT -Uri $uri -Headers (Build-AuthHeader $tok) -Body $Bytes -ContentType "application/octet-stream"
        return $u.id
    }

    $createUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$(Encode-RelForGraph $rel)`:/createUploadSession"
    $sess = Invoke-GraphJson -Method "POST" -Uri $createUri -Connexion $Connexion -Body @{ item = @{ "@microsoft.graph.conflictBehavior" = "replace" } }
    $uploadUrl = $sess.uploadUrl

    $total = $Bytes.Length
    $pos = 0
    $lastResp = $null

    while ($pos -lt $total) {
        $chunk = [Math]::Min($ChunkSize, $total - $pos)
        $from = $pos
        $to = $pos + $chunk - 1

        $content = [System.Net.Http.ByteArrayContent]::new($Bytes, $from, $chunk)
        $content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::new("application/octet-stream")
        $null = $content.Headers.TryAddWithoutValidation("Content-Range", ("bytes {0}-{1}/{2}" -f $from, $to, $total))

        $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Put, $uploadUrl)
        $req.Content = $content
        $resp = $script:HttpClient.SendAsync($req).Result

        if (-not $resp.IsSuccessStatusCode) {
            throw "Upload HTTP $([int]$resp.StatusCode)"
        }

        $txt = $resp.Content.ReadAsStringAsync().Result
        if ($txt) {
            try { $lastResp = $txt | ConvertFrom-Json } catch {}
        }

        $pos += $chunk
    }

    if ($lastResp -and $lastResp.id) { return $lastResp.id }
    $check = Item-Exists -DriveId $DriveId -Rel $rel -AccessToken $tok
    return $check.Id
}

function Load-UrlTranslations {
    if (-not (Test-Path $translation_csv)) {
        Write-Log "Fichier de traduction introuvable: $translation_csv" "WARN"
        return
    }

    $csv = Import-Csv -Path $translation_csv -Encoding UTF8
    foreach ($r in $csv) {
        if ($r.Source -and $r.Destination) {
            $script:UrlTranslations[(Normalize-RelPath $r.Source)] = (Normalize-RelPath $r.Destination)
        }
    }
    Write-Log ("Traductions chargees: {0}" -f $script:UrlTranslations.Count) "SUCCESS"
}

function Apply-UrlTranslations {
    param([string]$Path)

    $p = Normalize-RelPath $Path
    foreach ($k in $script:UrlTranslations.Keys) {
        if ($p.Contains($k)) {
            $p = $p.Replace($k, $script:UrlTranslations[$k])
        }
    }
    return $p
}

function Resolve-DestinationObjectId {
    param([string]$DisplayName, [string]$TargetType)

    $cacheKey = "$TargetType|$DisplayName"
    if ($script:ResolvedIdCache.ContainsKey($cacheKey)) { return $script:ResolvedIdCache[$cacheKey] }

    $safe = $DisplayName.Replace("'","''")
    $isGroup = $TargetType.Trim().ToLowerInvariant() -in @("group","securitygroup","m365group","microsoft365group","aadgroup")

    $uri = if ($isGroup) {
        "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe'&`$select=id&`$top=1"
    } else {
        "https://graph.microsoft.com/v1.0/users?`$filter=displayName eq '$safe'&`$select=id&`$top=1"
    }

    try {
        $r = Invoke-GraphJson -Method "GET" -Uri $uri -Connexion $connexion_cible
        $obj = $r.value | Select-Object -First 1
        $id = if ($obj) { $obj.id } else { $null }
        $script:ResolvedIdCache[$cacheKey] = $id
        if (-not $id) { Write-Log ("Resolution destination introuvable: {0}" -f $DisplayName) "WARN" }
        return $id
    } catch {
        $script:ResolvedIdCache[$cacheKey] = $null
        return $null
    }
}

function Apply-Item-Permissions {
    param([string]$PathOriginal, [string]$PathTranslated, [string]$ItemType, [string]$DstDriveId, [string]$ItemId, [object[]]$PermRows)

    if (-not $PermRows -or $PermRows.Count -eq 0) { return }
    $applied = 0
    $skipped = 0
    $failed = 0

    foreach ($perm in $PermRows) {
        $srcId = if ($perm.grantedToID) { [string]$perm.grantedToID } else { "" }
        $srcName = if ($perm.GrantedTo) { [string]$perm.GrantedTo } else { "" }
        $srcRole = if ($perm.Role) { [string]$perm.Role } else { "read" }
        $targetType = if ($perm.TargetType) { [string]$perm.TargetType } else { "User" }

        $entry = $null
        if ($srcId -and $script:UserMapping.ContainsKey($srcId)) {
            $entry = $script:UserMapping[$srcId]
        } elseif ($srcName -and $script:UserMappingByName.ContainsKey($srcName)) {
            $entry = $script:UserMappingByName[$srcName]
        }

        if (-not $entry) {
            Add-PermissionResult -PathOriginal $PathOriginal -PathTranslated $PathTranslated -ItemType $ItemType -GrantedTo $srcName -GrantedToId $srcId -TargetType $targetType -Role $srcRole -Status "skipped" -ErrorMessage "Mapping manquant" -MappedDestination $null -MappedDestinationId $null
            $skipped++
            continue
        }

        $dstId = Resolve-DestinationObjectId -DisplayName $entry.DisplayName -TargetType $entry.TargetType
        if (-not $dstId) {
            Add-PermissionResult -PathOriginal $PathOriginal -PathTranslated $PathTranslated -ItemType $ItemType -GrantedTo $srcName -GrantedToId $srcId -TargetType $targetType -Role $srcRole -Status "failed" -ErrorMessage "Resolution destination impossible" -MappedDestination $entry.DisplayName -MappedDestinationId $null
            $failed++
            continue
        }

        $role = switch ($srcRole.ToLowerInvariant()) {
            "owner" { "write" }
            "write" { "write" }
            default { "read" }
        }

        $inviteUri = "https://graph.microsoft.com/v1.0/drives/$DstDriveId/items/$ItemId/invite"
        $body = @{
            requireSignIn = $true
            sendInvitation = $false
            roles = @($role)
            recipients = @(@{ objectId = $dstId })
        }

        try {
            $null = Invoke-GraphJson -Method "POST" -Uri $inviteUri -Connexion $connexion_cible -Body $body
            Add-PermissionResult -PathOriginal $PathOriginal -PathTranslated $PathTranslated -ItemType $ItemType -GrantedTo $srcName -GrantedToId $srcId -TargetType $targetType -Role $role -Status "applied" -ErrorMessage $null -MappedDestination $entry.DisplayName -MappedDestinationId $dstId
            $applied++
        } catch {
            $msg = $_.Exception.Message
            Add-PermissionResult -PathOriginal $PathOriginal -PathTranslated $PathTranslated -ItemType $ItemType -GrantedTo $srcName -GrantedToId $srcId -TargetType $targetType -Role $role -Status "failed" -ErrorMessage $msg -MappedDestination $entry.DisplayName -MappedDestinationId $dstId
            Add-ErrorRow -Path $PathOriginal -Phase "permissions" -Message $msg
            $failed++
        }
    }

    Write-Log ("Permissions {0}: applied={1}, skipped={2}, failed={3}" -f $PathOriginal, $applied, $skipped, $failed)
}

function Copy-One {
    param([string]$Rel, [string]$ItemType, [string]$SrcDriveId, [string]$DstDriveId, [string]$SrcToken, [string]$DstToken)

    $originalRel = Normalize-RelPath (Strip-LibraryPrefix (Normalize-RelPath $Rel))
    $translatedRel = Normalize-RelPath (Apply-UrlTranslations $originalRel)

    if ([string]::IsNullOrWhiteSpace($originalRel)) {
        return @{ Ok = $false; ItemId = $null; TranslatedRel = $translatedRel; IsFolder = $false; ItemType = $ItemType }
    }

    $dstExisting = Item-Exists -DriveId $DstDriveId -Rel $translatedRel -AccessToken $DstToken
    if ($dstExisting.Exists -and -not $ForceOverwrite) {
        Write-Log ("Element deja present, overwrite desactive: {0}" -f $translatedRel)
        return @{ Ok = $true; ItemId = $dstExisting.Id; TranslatedRel = $translatedRel; IsFolder = ($ItemType -eq "Folder"); ItemType = $ItemType }
    }

    $srcExisting = Item-Exists -DriveId $SrcDriveId -Rel $originalRel -AccessToken $SrcToken
    if (-not $srcExisting.Exists) {
        Add-ErrorRow -Path $originalRel -Phase "copy" -Message "Source introuvable"
        Write-Log ("Source introuvable: {0}" -f $originalRel) "WARN"
        return @{ Ok = $false; ItemId = $null; TranslatedRel = $translatedRel; IsFolder = $false; ItemType = $ItemType }
    }

    if ($srcExisting.IsFolder -or $ItemType -eq "Folder") {
        $folderId = Ensure-FolderPath -DriveId $DstDriveId -FolderRel $translatedRel -AccessToken $DstToken
        return @{ Ok = $true; ItemId = $folderId; TranslatedRel = $translatedRel; IsFolder = $true; ItemType = "Folder" }
    }

    $bytes = Download-FileContent -DriveId $SrcDriveId -RelativePath $originalRel -Connexion $connexion_source
    $newItemId = Upload-File -DriveId $DstDriveId -RelativePath $translatedRel -Bytes $bytes -Connexion $connexion_cible

    if (-not $newItemId) {
        Add-ErrorRow -Path $originalRel -Phase "copy" -Message "Upload termine sans item id"
        Write-Log ("Upload sans item id: {0}" -f $translatedRel) "WARN"
        return @{ Ok = $false; ItemId = $null; TranslatedRel = $translatedRel; IsFolder = $false; ItemType = "File" }
    }

    return @{ Ok = $true; ItemId = $newItemId; TranslatedRel = $translatedRel; IsFolder = $false; ItemType = "File" }
}

# =========================
# LISTES / COLONNES / ELEMENTS
# =========================
function Test-CustomColumn($Col) {
    if ($Col.readOnly -or $Col.hidden) { return $false }
    if ($FieldSkipNames -contains [string]$Col.name) { return $false }
    if ($Col.columnGroup -and ($NativeColumnGroups -contains [string]$Col.columnGroup)) { return $false }
    return $true
}

function Convert-ColumnForCreate {
    param($Col, [string]$TargetListIdForLookup)

    $o = [ordered]@{ name = [string]$Col.name; displayName = [string]$Col.displayName }
    if ($Col.description) { $o["description"] = [string]$Col.description }
    if ($Col.required) { $o["required"] = $true }
    if ($Col.indexed) { $o["indexed"] = $true }
    if ($Col.enforceUniqueValues) { $o["enforceUniqueValues"] = $true }

    $facet = $null
    foreach ($f in @('text','choice','number','dateTime','boolean','currency','personOrGroup','hyperlinkOrPicture','calculated','lookup','geolocation','term','thumbnail','contentApprovalStatus')) {
        if ($Col.PSObject.Properties[$f] -and $null -ne $Col.$f) { $facet = $f; break }
    }
    if (-not $facet) { return $null }

    if ($facet -eq 'lookup') {
        if (-not $TargetListIdForLookup) { return $null }
        $srcLookup = $Col.lookup
        $o["lookup"] = @{ listId = $TargetListIdForLookup; columnName = [string]$srcLookup.columnName; allowMultipleValues = [bool]$srcLookup.allowMultipleValues; allowUnlimitedLength = [bool]$srcLookup.allowUnlimitedLength }
    } else {
        $o[$facet] = $Col.$facet
    }
    return $o
}

function Get-TargetColumns {
    param([string]$DstSiteId, [string]$TargetListId)
    return @(Get-GraphAllPages -Uri "https://graph.microsoft.com/v1.0/sites/$DstSiteId/lists/$TargetListId/columns" -Connexion $connexion_cible)
}

function Update-WritableColumns {
    param([string]$DstSiteId, [string]$TargetListId)

    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($c in (Get-TargetColumns -DstSiteId $DstSiteId -TargetListId $TargetListId)) {
        if (-not $c.readOnly -and -not $c.hidden -and ($FieldSkipNames -notcontains [string]$c.name)) { [void]$set.Add([string]$c.name) }
    }
    $script:WritableColumns[$TargetListId] = $set
}

# Colonnes de la liste source -> colonnes de la cible (hors colonnes natives ; lookups traites dans une 2e passe).
function Sync-ListColumns {
    param($SrcList, [string]$DstSiteId, [string]$TargetListId, [bool]$LookupPass)

    $srcCols = @($SrcList.columns)
    if ($srcCols.Count -eq 0) { return }
    $existing = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($c in (Get-TargetColumns -DstSiteId $DstSiteId -TargetListId $TargetListId)) { [void]$existing.Add([string]$c.name) }

    foreach ($col in $srcCols) {
        if (-not (Test-CustomColumn $col)) { continue }
        $isLookup = ($col.PSObject.Properties['lookup'] -and $null -ne $col.lookup)
        if ($isLookup -ne $LookupPass) { continue }
        if ($existing.Contains([string]$col.name)) { continue }

        $lookupTarget = $null
        if ($isLookup) {
            $lookupTarget = $script:ListIdMap[[string]$col.lookup.listId]
            if (-not $lookupTarget) {
                Write-Log ("Colonne lookup '{0}' ignoree (liste referencee non copiee)" -f $col.name) "WARN"
                continue
            }
        }

        $body = Convert-ColumnForCreate -Col $col -TargetListIdForLookup $lookupTarget
        if (-not $body) { continue }
        try {
            $null = Invoke-GraphJson -Method "POST" -Uri "https://graph.microsoft.com/v1.0/sites/$DstSiteId/lists/$TargetListId/columns" -Connexion $connexion_cible -Body $body
            Write-Log ("Colonne creee: {0} / {1}" -f $SrcList.relativeUrl, $col.name)
        } catch {
            Write-Log ("Colonne '{0}' non creee sur '{1}': {2}" -f $col.name, $SrcList.relativeUrl, $_.Exception.Message) "WARN"
        }
    }
}

# Retrouve ou cree la liste/bibliotheque cible. Retourne @{ ListId; Key; IsLibrary; DriveId } ou $null.
function Ensure-TargetList {
    param($SrcList, [string]$DstSiteId, [hashtable]$DstLists)

    $targetKey = Get-TargetKey ([string]$SrcList.relativeUrl)
    $lk = $targetKey.ToLowerInvariant()
    if ($DstLists.ContainsKey($lk)) { return $DstLists[$lk] }
    if (-not $CreateMissingLists) { return $null }

    $segment = ($targetKey -split '/')[-1]
    $isLib = [bool]$SrcList.isLibrary
    $fallback = if ($isLib) { "documentLibrary" } else { "genericList" }
    $template = if ($SrcList.template) { [string]$SrcList.template } else { $fallback }

    # displayName = segment d'URL a la creation (fixe l'URL), puis renommage vers le vrai titre.
    $created = $null
    foreach ($tpl in @($template, $fallback) | Select-Object -Unique) {
        try {
            $created = Invoke-GraphJson -Method "POST" -Uri "https://graph.microsoft.com/v1.0/sites/$DstSiteId/lists" -Connexion $connexion_cible -Body @{ displayName = $segment; list = @{ template = $tpl } }
            break
        } catch {
            Write-Log ("Creation de '{0}' avec le modele '{1}' impossible: {2}" -f $segment, $tpl, $_.Exception.Message) "WARN"
        }
    }
    if (-not $created) { throw "Creation de la liste '$targetKey' impossible" }

    if ($SrcList.displayName -and $SrcList.displayName -ne $segment) {
        try {
            $null = Invoke-GraphJson -Method "PATCH" -Uri "https://graph.microsoft.com/v1.0/sites/$DstSiteId/lists/$($created.id)" -Connexion $connexion_cible -Body @{ displayName = [string]$SrcList.displayName }
        } catch {
            Write-Log ("Renommage de '{0}' impossible: {1}" -f $segment, $_.Exception.Message) "WARN"
        }
    }

    $driveId = $null
    if ($isLib) {
        $d = Invoke-GraphJson -Method "GET" -Uri "https://graph.microsoft.com/v1.0/sites/$DstSiteId/lists/$($created.id)/drive?`$select=id" -Connexion $connexion_cible
        $driveId = $d.id
    }

    $entry = @{ ListId = [string]$created.id; Key = $targetKey; IsLibrary = $isLib; DriveId = $driveId }
    $DstLists[$lk] = $entry
    Write-Log ("Liste/bibliotheque creee sur la cible: {0}" -f $targetKey) "SUCCESS"
    return $entry
}

function ConvertTo-WritableFields {
    param($Fields, [string]$TargetListId)

    $payload = @{}
    if (-not $Fields) { return $payload }
    $writable = $script:WritableColumns[$TargetListId]

    foreach ($prop in $Fields.PSObject.Properties) {
        $name = [string]$prop.Name
        if ($name.StartsWith('@') -or $name -match '@odata') { continue }
        if ($name -match 'LookupId$|LookupValue$') { continue }   # personnes / lookups : IDs propres a la source
        if ($FieldSkipNames -contains $name) { continue }
        if ($writable -and -not $writable.Contains($name)) { continue }
        if ($null -eq $prop.Value) { continue }
        $payload[$name] = $prop.Value
        if ($prop.Value -is [System.Array] -and $prop.Value.Count -gt 0 -and $prop.Value[0] -is [string]) {
            $payload["$name@odata.type"] = "Collection(Edm.String)"
        }
    }
    return $payload
}

function Copy-ListItems {
    param($SrcList, [string]$DstSiteId, [string]$TargetListId)

    $file = Join-Path $DiscoveryDirectory ("listitems\{0}.jsonl" -f $SrcList.listId)
    $rows = @(Read-Jsonl -Path $file)
    Write-Log ("Elements a copier pour '{0}': {1}" -f $SrcList.relativeUrl, $rows.Count)
    if ($rows.Count -eq 0) { return }
    $rows = @($rows | Sort-Object { [int]$_.id })

    $itemsUri = "https://graph.microsoft.com/v1.0/sites/$DstSiteId/lists/$TargetListId/items"
    $existing = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($e in (Get-GraphAllPages -Uri ($itemsUri + '?$select=id&$top=500') -Connexion $connexion_cible)) { [void]$existing.Add([int]$e.id) }
    $next = 1
    if ($existing.Count -gt 0) { $next = ([int]($existing | Measure-Object -Maximum).Maximum) + 1 }

    $placeholders = New-Object System.Collections.Generic.List[int]
    $placeholderBudget = $MaxIdPlaceholders
    $idWarned = $false

    foreach ($row in $rows) {
        $srcId = [int]$row.id
        $path = "$($SrcList.relativeUrl)/ID:$srcId"
        try {
            $payload = if ($CopyItemFields) { ConvertTo-WritableFields -Fields $row.fields -TargetListId $TargetListId } else { @{} }

            if ($existing.Contains($srcId)) {
                if ($ForceOverwrite -and $payload.Count -gt 0) {
                    $null = Invoke-GraphJson -Method "PATCH" -Uri "$itemsUri/$srcId/fields" -Connexion $connexion_cible -Body $payload
                }
                Add-ListItemResult -ListKey $SrcList.relativeUrl -SourceItemId $srcId -Status "copied" -TargetItemId $srcId -ErrorMessage $null
                continue
            }

            # Conserve les ID : combler les trous par des elements factices supprimes a la fin.
            while ($next -lt $srcId -and $placeholderBudget -gt 0) {
                $ph = $null
                try { $ph = Invoke-GraphJson -Method "POST" -Uri $itemsUri -Connexion $connexion_cible -Body @{ fields = @{ Title = "__placeholder__" } } } catch {
                    $ph = Invoke-GraphJson -Method "POST" -Uri $itemsUri -Connexion $connexion_cible -Body @{ fields = @{} }
                }
                $placeholders.Add([int]$ph.id)
                $next = [int]$ph.id + 1
                $placeholderBudget--
            }

            $newItem = Invoke-GraphJson -Method "POST" -Uri $itemsUri -Connexion $connexion_cible -Body @{ fields = $payload }
            $newId = [int]$newItem.id
            $next = $newId + 1
            if ($newId -ne $srcId -and -not $idWarned) {
                $idWarned = $true
                Write-Log ("ID non preserves sur '{0}' (source {1} -> cible {2}) : droitscustome_2.ps1 ne retrouvera pas ces elements" -f $SrcList.relativeUrl, $srcId, $newId) "WARN"
            }
            Add-ListItemResult -ListKey $SrcList.relativeUrl -SourceItemId $srcId -Status "copied" -TargetItemId $newId -ErrorMessage $null
        } catch {
            $msg = $_.Exception.Message
            Write-Log ("Element {0} non copie: {1}" -f $path, $msg) "ERROR"
            Add-ListItemResult -ListKey $SrcList.relativeUrl -SourceItemId $srcId -Status "failed" -TargetItemId $null -ErrorMessage $msg
            Add-ErrorRow -Path $path -Phase "list-items" -Message $msg
        }
    }

    foreach ($phId in $placeholders) {
        try { $null = Invoke-GraphJson -Method "DELETE" -Uri "$itemsUri/$phId" -Connexion $connexion_cible } catch {
            Write-Log ("Element factice {0} non supprime sur '{1}': {2}" -f $phId, $SrcList.relativeUrl, $_.Exception.Message) "WARN"
        }
    }
    if ($placeholders.Count -gt 0) {
        Write-Log ("{0} element(s) factice(s) supprime(s) sur '{1}' (conservation des ID)" -f $placeholders.Count, $SrcList.relativeUrl)
    }
}

function Copy-ItemFieldsToTarget {
    param($Fields, [string]$OriginalPath, [string]$DstDriveId, [string]$DstItemId, [string]$TargetListId)

    $payload = ConvertTo-WritableFields -Fields $Fields -TargetListId $TargetListId
    if ($payload.Count -eq 0) { return }
    try {
        $null = Invoke-GraphJson -Method "PATCH" -Uri "https://graph.microsoft.com/v1.0/drives/$DstDriveId/items/$DstItemId/listItem/fields" -Connexion $connexion_cible -Body $payload
    } catch {
        Write-Log ("Colonnes non copiees pour {0}: {1}" -f $OriginalPath, $_.Exception.Message) "WARN"
    }
}

# =========================
# MAIN
# =========================
try {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    $script:LogFile = Join-Path $OutputDirectory 'copie.log'
    Set-Content -LiteralPath $script:LogFile -Value '' -Encoding UTF8
    Write-Log "Demarrage copie (decouverte: $DiscoveryDirectory, sortie: $OutputDirectory)"

    $listsFile = Join-Path $DiscoveryDirectory 'lists.json'
    if (-not (Test-Path -LiteralPath $listsFile)) {
        throw "lists.json introuvable dans $DiscoveryDirectory : lancez d'abord DecouverteDesFichers.ps1"
    }
    $parsedLists = Get-Content -LiteralPath $listsFile -Raw -Encoding UTF8 | ConvertFrom-Json
    $discoveredLists = @($parsedLists | ForEach-Object { $_ })   # aplatit le tableau (PowerShell 5.1)
    if ($discoveredLists.Count -eq 0) { throw "lists.json vide" }
    Write-Log ("Listes et bibliotheques a traiter: {0}" -f $discoveredLists.Count)

    Load-UrlTranslations
    Load-UserMapping

    $srcCtx = Get-SiteContext -SiteUrl $site_url_source -Connexion $connexion_source
    $dstCtx = Get-SiteContext -SiteUrl $site_url_cible -Connexion $connexion_cible
    $dstSiteId = $dstCtx.Site.id

    # ---- Etape 1 : listes/bibliotheques et colonnes sur la cible ----
    $dstByListId = @{}   # id liste source -> entree cible
    foreach ($sl in $discoveredLists) {
        try {
            $entry = Ensure-TargetList -SrcList $sl -DstSiteId $dstSiteId -DstLists $dstCtx.Lists
            if (-not $entry) {
                Write-Log ("Liste absente de la cible et creation desactivee: {0}" -f $sl.relativeUrl) "WARN"
                Add-ListResult -ListKey $sl.relativeUrl -Status "failed" -TargetListId $null -ErrorMessage "Absente de la cible"
                continue
            }
            $dstByListId[[string]$sl.listId] = $entry
            $script:ListIdMap[[string]$sl.listId] = $entry.ListId
            Sync-ListColumns -SrcList $sl -DstSiteId $dstSiteId -TargetListId $entry.ListId -LookupPass $false
        } catch {
            $msg = $_.Exception.Message
            Write-Log ("Liste '{0}' non preparee: {1}" -f $sl.relativeUrl, $msg) "ERROR"
            Add-ListResult -ListKey $sl.relativeUrl -Status "failed" -TargetListId $null -ErrorMessage $msg
            Add-ErrorRow -Path $sl.relativeUrl -Phase "lists" -Message $msg
        }
    }
    # Colonnes lookup : toutes les listes existent maintenant, les ID de listes peuvent etre transposes.
    foreach ($sl in $discoveredLists) {
        $entry = $dstByListId[[string]$sl.listId]
        if (-not $entry) { continue }
        try { Sync-ListColumns -SrcList $sl -DstSiteId $dstSiteId -TargetListId $entry.ListId -LookupPass $true } catch {
            Write-Log ("Colonnes lookup de '{0}': {1}" -f $sl.relativeUrl, $_.Exception.Message) "WARN"
        }
        try { Update-WritableColumns -DstSiteId $dstSiteId -TargetListId $entry.ListId } catch {}
    }

    # Index par bibliotheque (cle relative source, minuscules) pour la copie des fichiers.
    $libBySrcKey = @{}
    foreach ($sl in $discoveredLists) {
        if (-not [bool]$sl.isLibrary) { continue }
        $entry = $dstByListId[[string]$sl.listId]
        $srcEntry = $srcCtx.Lists[([string]$sl.relativeUrl).ToLowerInvariant()]
        if ($entry -and $srcEntry -and $srcEntry.DriveId -and $entry.DriveId) {
            $libBySrcKey[([string]$sl.relativeUrl).ToLowerInvariant()] = @{
                SrcDriveId = $srcEntry.DriveId; DstDriveId = $entry.DriveId; DstKey = $entry.Key; DstListId = $entry.ListId
            }
            Add-ListResult -ListKey $sl.relativeUrl -Status "prepared" -TargetListId $entry.ListId -ErrorMessage $null
        } else {
            Write-Log ("Bibliotheque ignoree (drive source/cible introuvable): {0}" -f $sl.relativeUrl) "WARN"
            Add-ListResult -ListKey $sl.relativeUrl -Status "failed" -TargetListId $null -ErrorMessage "Drive source/cible introuvable"
        }
    }

    # ---- Etape 2 : fichiers et dossiers des bibliotheques + droits ----
    $permsByPath = @{}
    foreach ($p in (Read-Jsonl -Path (Join-Path $DiscoveryDirectory 'permissions.jsonl'))) {
        if (-not $permsByPath.ContainsKey($p.path)) { $permsByPath[$p.path] = New-Object System.Collections.Generic.List[object] }
        $permsByPath[$p.path].Add([pscustomobject]@{
            GrantedTo = $p.grantedTo; grantedToID = $p.grantedToId; TargetType = $p.targetType; Role = $p.role
        })
    }
    Write-Log ("Chemins avec droits charges: {0}" -f $permsByPath.Count)

    $count = 0
    foreach ($row in (Read-Jsonl -Path (Join-Path $DiscoveryDirectory 'items.jsonl'))) {
        $count++
        $path = [string]$row.path
        $itemType = [string]$row.itemType
        $permRows = if ($permsByPath.ContainsKey($path)) { @($permsByPath[$path]) } else { @() }

        Write-Log ("Traitement [{0}]: {1}" -f $count, $path)

        try {
            $split = Split-LibraryPath $path
            $lib = $libBySrcKey[$split.Library.ToLowerInvariant()]
            if (-not $lib) {
                Add-ItemResult -PathOriginal $path -PathTranslated $null -ItemType $itemType -ItemId $null -Status "failed"
                Add-ErrorRow -Path $path -Phase "copy" -Message "Bibliotheque source/cible introuvable: $($split.Library)"
                continue
            }

            # Jetons rafraichis a chaque element (le cache renouvelle avant expiration).
            $srcToken = Get-AccessToken -Conn $connexion_source
            $dstToken = Get-AccessToken -Conn $connexion_cible

            $copy = Copy-One -Rel $split.Rel -ItemType $itemType -SrcDriveId $lib.SrcDriveId -DstDriveId $lib.DstDriveId -SrcToken $srcToken -DstToken $dstToken
            $translated = if ($copy.TranslatedRel) { "$($lib.DstKey)/$($copy.TranslatedRel)" } else { "$($lib.DstKey)/" }

            if ($copy.Ok) {
                Add-ItemResult -PathOriginal $path -PathTranslated $translated -ItemType $copy.ItemType -ItemId $copy.ItemId -Status "copied"
                if ($CopyItemFields -and $copy.ItemId -and $row.fields) {
                    Copy-ItemFieldsToTarget -Fields $row.fields -OriginalPath $path -DstDriveId $lib.DstDriveId -DstItemId $copy.ItemId -TargetListId $lib.DstListId
                }
                Apply-Item-Permissions -PathOriginal $path -PathTranslated $translated -ItemType $copy.ItemType -DstDriveId $lib.DstDriveId -ItemId $copy.ItemId -PermRows $permRows
                Write-Log ("Copie OK: {0}" -f $translated) "SUCCESS"
            } else {
                Add-ItemResult -PathOriginal $path -PathTranslated $translated -ItemType $itemType -ItemId $null -Status "failed"
                Write-Log ("Copie KO: {0}" -f $path) "WARN"
            }
        } catch {
            $msg = $_.Exception.Message
            Add-ItemResult -PathOriginal $path -PathTranslated $null -ItemType $itemType -ItemId $null -Status "failed"
            Add-ErrorRow -Path $path -Phase "copy" -Message $msg
            Write-Log ("Erreur: {0}" -f $msg) "ERROR"
        }
    }

    # ---- Etape 3 : elements des listes classiques ----
    if ($CopyListItems) {
        foreach ($sl in $discoveredLists) {
            if ([bool]$sl.isLibrary) { continue }
            $entry = $dstByListId[[string]$sl.listId]
            if (-not $entry) { continue }
            try {
                Copy-ListItems -SrcList $sl -DstSiteId $dstSiteId -TargetListId $entry.ListId
                Add-ListResult -ListKey $sl.relativeUrl -Status "copied" -TargetListId $entry.ListId -ErrorMessage $null
            } catch {
                $msg = $_.Exception.Message
                Write-Log ("Elements de '{0}' non copies: {1}" -f $sl.relativeUrl, $msg) "ERROR"
                Add-ListResult -ListKey $sl.relativeUrl -Status "failed" -TargetListId $entry.ListId -ErrorMessage $msg
                Add-ErrorRow -Path $sl.relativeUrl -Phase "list-items" -Message $msg
            }
        }
    }

    Write-Log "Copie terminee" "SUCCESS"
}
catch {
    $msg = $_.Exception.Message
    Write-Log ("Erreur fatale: {0}" -f $msg) "ERROR"
    Add-ErrorRow -Path "/" -Phase "fatal" -Message $msg
    $script:ExitCode = 1
}
finally {
    try { Save-Results } catch { Write-Host "Echec ecriture des resultats: $($_.Exception.Message)" }
    if ($script:HttpClient) { $script:HttpClient.Dispose() }
}

if ((Get-Variable -Name ExitCode -Scope Script -ErrorAction SilentlyContinue) -and $script:ExitCode -ne 0) {
    exit $script:ExitCode
}
