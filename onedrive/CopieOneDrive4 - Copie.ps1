# =========================
# CONFIGURATION
# =========================

# App Registration TENANT SOURCE
$app_source = @{
tenant_Id = "85a5a352-25d6-4894-a97d-221cd1712dd2"
        client_Id = "684f670e-6931-4120-b053-f8c9538744f6"
        client_Secret = "O5s8Q~Z~e_d1-dSQ.vt4uFTeR8c4Xssaa093Racz"
}

# App Registration TENANT CIBLE
$app_cible = @{
    tenant_id     = "2eea08b8-1972-447b-ad43-d044d042500a"
    client_id     = " 821109a4-6e0e-48e8-b477-8f9b70aa32a4"
    client_secret = "Yhy8Q~mBxTrYqx~ve7enqTQatjwhEZ5.34-q-a4j"
}


# =========================
# CHEMIN DU FICHIER CSV
# Format attendu du CSV (séparateur ;) :
#   onedrive_url_source;onedrive_url_cible;upn_source;upn_cible
# Exemple :
#   https://ipdlab-my.sharepoint.com/personal/user01_ipd-lab_com/_layouts/15/onedrive.aspx;https://infoprodigital365-my.sharepoint.com/personal/testmig01_infopro-digital_com/_layouts/15/onedrive.aspx;user01@ipd-lab.com;testmig01@infopro-digital.fr
# =========================
$CsvPath = "$PSScriptRoot\migrations.csv"

# Options
$ForceOverwrite = $true
$ChunkSizeMB    = 10
$CopyVersionHistory = $true
$MaxVersionsPerFile = 0
$PreserveFileTimestamps = $true
$ExportVersionMetadata  = $true

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
    $url = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/items/" + $ItemId + "/versions?" + '$top=200&$select=id,lastModifiedDateTime,size'

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
    $segs    = ($RelPath -replace '\\', '/') -split '/' | Where-Object { $_ } |
               ForEach-Object { [Uri]::EscapeDataString($_) }
    $encPath = $segs -join '/'
    $uri     = "https://graph.microsoft.com/v1.0/drives/" + $DriveId + "/root:/" + $encPath + ":/content"
    Invoke-GraphCall -Uri $uri -Token $Token -Method "PUT" -ByteBody $Bytes `
                     -ContentType "application/octet-stream" | Out-Null
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
                Write-Log "INFO" "   ⬆  Rejeu de $($versions.Count) version(s) : $($Item.RelPath)"
                $vCounter = 0
                foreach ($v in $versions) {
                    $vCounter++
                    $versionId = ""
                    $versionWhen = ""
                    $versionBy = ""
                    try { if ($v.id) { $versionId = [string]$v.id } } catch {}
                    try { if ($v.lastModifiedDateTime) { $versionWhen = [string]$v.lastModifiedDateTime } } catch {}
                    try {
                        if ($v.lastModifiedBy.user.email) { $versionBy = [string]$v.lastModifiedBy.user.email }
                        elseif ($v.lastModifiedBy.user.displayName) { $versionBy = [string]$v.lastModifiedBy.user.displayName }
                        elseif ($v.lastModifiedBy.application.displayName) { $versionBy = [string]$v.lastModifiedBy.application.displayName }
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
                    Write-Log "DEBUG" "      Version $vCounter/$($versions.Count) rejouee (id=$versionId, $($vBytes.Length) o)"
                }

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

        if ($CopyVersionHistory) {
            Write-Log "WARN" "Graph ne permet pas de forcer l'auteur des versions. La version cible est attribuee au compte d'execution."
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

    # Extraire le folder depuis les URLs
    $srcFolder = (([Uri]$OneDriveUrlSource).AbsolutePath -split '/')[2]
    $dstFolder = (([Uri]$OneDriveUrlCible).AbsolutePath  -split '/')[2]

    # Tokens
    $srcToken = Get-GraphToken -TenantId $app_source.tenant_id `
                -ClientId $app_source.client_id -ClientSecret $app_source.client_secret
    $dstToken = Get-GraphToken -TenantId $app_cible.tenant_id `
                -ClientId $app_cible.client_id  -ClientSecret $app_cible.client_secret

    # Résolution des Drive IDs
    Write-Log "INFO" "=== RESOLUTION DES DRIVES ==="
    $srcDriveId = Get-DriveId -UPN $UpnSource -UserFolder $srcFolder -Token $srcToken -Label "SOURCE"
    $dstDriveId = Get-DriveId -UPN $UpnCible  -UserFolder $dstFolder -Token $dstToken -Label "CIBLE"
    Write-Log "SUCCESS" "Drive source : $srcDriveId"
    Write-Log "SUCCESS" "Drive cible  : $dstDriveId"

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
        $dstToken = Get-GraphToken -TenantId $app_cible.tenant_id `
                    -ClientId $app_cible.client_id -ClientSecret $app_cible.client_secret
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
        $dstToken = Get-GraphToken -TenantId $app_cible.tenant_id `
                    -ClientId $app_cible.client_id  -ClientSecret $app_cible.client_secret
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