# Script PowerShell - Decouverte SharePoint (toutes les listes et bibliotheques) -> fichiers
# Version: 2.0 (sans base de donnees)
#
# Sorties dans $OutputDirectory (lues ensuite par copieDesFichiers3-historique2.ps1) :
#   lists.json             listes/bibliotheques + definitions de colonnes
#   items.jsonl            fichiers/dossiers des bibliotheques (+ valeurs des colonnes)
#   permissions.jsonl      droits Graph de chaque fichier/dossier
#   listitems\<id>.jsonl   elements des listes classiques
#   discovery.log          journal
# + $exportPath (CSV des permissions) et $mappingPath (UserMapping.csv a completer).

try {
    if ($Host -and $Host.Name -and $Host.Name -notmatch 'ISE') {
        if (-not [Console]::IsInputRedirected) { [Console]::InputEncoding = [System.Text.Encoding]::UTF8 }
        if (-not [Console]::IsOutputRedirected) { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 }
    }
} catch {}
try { if ($Host -and $Host.UI) { $OutputEncoding = [System.Text.Encoding]::UTF8 } } catch {}

# =========================
# CONFIG GRAPH
# =========================
$tenantId     = "2eea08b8-1972-447b-ad43-d044d042500a"
$clientId     = "821109a4-6e0e-48e8-b477-8f9b70aa32a4"
    $clientSecret = "XYr8Q~jx8vDX30c97YP5M0-juBIahzni6WlDNdz8"
$siteUrl = "https://graph.microsoft.com/v1.0/sites/infoprodigital365.sharepoint.com/:/sites/fethygate"

# Perimetre : TOUTES les bibliotheques et listes du site.
# Vide = tout le site. Sinon noms (titre ou segment d'URL, ex. "Documents", "Shared Documents", "Lists/Taches").
$ListNameFilter = @()
$IncludeHiddenLists = $false
# Chemins relatifs (jokers autorises) toujours ignores.
$ExcludedListUrlPatterns = @('appdata', 'fpdatasources', 'FormServerTemplates', 'Style Library', '_catalogs/*', 'SiteCollectionImages', 'Lists/DO_NOT_DELETE*')
# $true = enregistre aussi les elements des listes classiques.
$ScanListItems = $true
# $true = enregistre les valeurs des colonnes (metadonnees) des fichiers/dossiers.
$CaptureItemFields = $true

# Dossier de sortie (relu par le script de copie).
$baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$OutputDirectory = Join-Path $baseDir 'SharePoint-Discovery'

# Exports CSV
$enableCsvExport = $true
$exportPath  = "E:\SharePoint_Permissions_Export_src9999.csv"
$mappingPath = "E:\UserMapping.csv"

# =========================
# ETAT
# =========================
$script:accessToken = $null
$script:tokenExpiry = $null
$script:SiteId = $null
$script:SiteWebUrl = $null
$script:CurrentLibrary = $null
$script:ExpandFieldsSupported = $true
$script:allPaths = [System.Collections.Generic.HashSet[string]]::new()
$script:permRowsForCsv = New-Object System.Collections.Generic.List[object]
$script:mappingRows = @{}
$script:Writers = @{}
$script:Counts = @{ items = 0; permissions = 0; listItems = 0; errors = 0 }
$script:LogFile = $null

# =========================
# UTILS
# =========================
function Write-Log {
    param([string]$message, [string]$level = "INFO")
    $line = "[{0}][{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $level, $message
    Write-Host $line
    if ($script:LogFile) {
        try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 } catch {}
    }
}

function Open-Jsonl {
    param([string]$Name)
    $path = Join-Path $OutputDirectory $Name
    $dir = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $w = New-Object System.IO.StreamWriter($path, $false, (New-Object System.Text.UTF8Encoding($false)))
    $w.AutoFlush = $true
    $script:Writers[$Name] = $w
    return $w
}

function Write-Jsonl {
    param([System.IO.StreamWriter]$Writer, $Object)
    $Writer.WriteLine(($Object | ConvertTo-Json -Depth 8 -Compress))
}

function Close-AllWriters {
    foreach ($w in $script:Writers.Values) { try { $w.Dispose() } catch {} }
    $script:Writers = @{}
}

# =========================
# AUTH / GRAPH
# =========================
function Get-AccessToken {
    param([string]$clientId, [string]$clientSecret, [string]$tenantId)

    if ($script:accessToken -and $script:tokenExpiry -and (Get-Date).AddMinutes(5) -lt $script:tokenExpiry) {
        return $script:accessToken
    }

    $tokenBody = @{
        grant_type    = "client_credentials"
        scope         = "https://graph.microsoft.com/.default"
        client_id     = $clientId
        client_secret = $clientSecret
    }

    $tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $tokenBody
    $script:accessToken = $tokenResponse.access_token
    $script:tokenExpiry = (Get-Date).AddSeconds($tokenResponse.expires_in)
    return $script:accessToken
}

function Invoke-With-Retry {
    param(
        [string]$Uri,
        [string]$Method = "GET",
        $Body = $null,
        [int]$MaxRetries = 5
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $token = Get-AccessToken -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId
            $headers = @{ Authorization = "Bearer $token" }

            if ($Method -eq "GET") {
                return Invoke-RestMethod -Method Get -Uri $Uri -Headers $headers
            }
            return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body ($Body | ConvertTo-Json -Depth 10) -ContentType "application/json"
        } catch {
            $statusCode = $null
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}

            if ($statusCode -eq 401) {
                $script:accessToken = $null
                $script:tokenExpiry = $null
            }

            $retryable = ($statusCode -in @(401,408,429)) -or ($statusCode -ge 500 -and $statusCode -lt 600)
            if ($retryable -and $attempt -lt $MaxRetries) {
                Start-Sleep -Seconds ([Math]::Min(20, [Math]::Pow(2, $attempt)))
                continue
            }
            throw
        }
    }
}

function Get-AllPages {
    param([string]$Uri)

    $all = New-Object System.Collections.Generic.List[object]
    while ($Uri) {
        $r = Invoke-With-Retry -Uri $Uri
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

function Test-ListSelected {
    param([string]$Key, [string]$DisplayName, [string]$Name, [bool]$Hidden)

    if ($Hidden -and -not $IncludeHiddenLists) { return $false }
    foreach ($pattern in $ExcludedListUrlPatterns) {
        if ($Key -like $pattern) { return $false }
    }
    if ($ListNameFilter.Count -eq 0) { return $true }
    foreach ($f in $ListNameFilter) {
        if ($Key -ieq $f -or $DisplayName -ieq $f -or $Name -ieq $f -or ($Key -split '/')[-1] -ieq $f) { return $true }
    }
    return $false
}

# =========================
# SCAN
# =========================
function Add-ErrorRow {
    param([string]$path, [string]$phase, [string]$message)
    $script:Counts.errors++
    Write-Jsonl $script:Writers['errors.jsonl'] ([ordered]@{ path = $path; phase = $phase; message = $message })
}

function Get-Item-Permissions {
    param(
        [string]$driveId,
        [string]$itemId,
        [string]$path,
        [string]$type
    )

    $permUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$itemId/permissions"
    $permCount = 0
    $roleCount = 0

    do {
        $permResponse = Invoke-With-Retry -Uri $permUrl
        if (-not $permResponse) { return }

        foreach ($perm in $permResponse.value) {
            $permCount++
            $grantedTo = "<unknown>"
            $targetType = "Unknown"
            $grantedToID = "Unknown"

            if ($perm.grantedToV2 -and $perm.grantedToV2.user) {
                $grantedTo = $perm.grantedToV2.user.displayName
                $targetType = "User"
                $grantedToID = $perm.grantedToV2.user.id
            } elseif ($perm.grantedToV2 -and $perm.grantedToV2.group) {
                $grantedTo = $perm.grantedToV2.group.displayName
                $targetType = "Group"
                $grantedToID = $perm.grantedToV2.group.id
            } elseif ($perm.grantedTo -and $perm.grantedTo.user) {
                $grantedTo = $perm.grantedTo.user.displayName
                $targetType = "User"
                $grantedToID = $perm.grantedTo.user.id
            }

            foreach ($role in $perm.roles) {
                $roleCount++
                $script:Counts.permissions++
                Write-Jsonl $script:Writers['permissions.jsonl'] ([ordered]@{
                    path = $path; itemType = $type; grantedTo = $grantedTo; grantedToId = $grantedToID; targetType = $targetType; role = $role
                })

                if ($grantedTo -and $grantedTo -ne '<unknown>' -and $grantedToID -ne 'Unknown' -and -not $script:mappingRows.ContainsKey($grantedTo)) {
                    $script:mappingRows[$grantedTo] = [pscustomobject]@{
                        SourceId = $grantedToID; SourceDisplayName = $grantedTo; TargetType = $targetType; DestinationDisplayName = ''
                    }
                }

                if ($enableCsvExport) {
                    $script:permRowsForCsv.Add([pscustomobject]@{
                        Path = $path; ItemType = $type; GrantedTo = $grantedTo; grantedToID = $grantedToID; TargetType = $targetType; Role = $role
                    }) | Out-Null
                }
            }
        }

        $permUrl = $permResponse.'@odata.nextLink'
    } while ($permUrl)

    Write-Log ("Permissions detectees: item={0}, perms={1}, roles={2}" -f $path, $permCount, $roleCount) "DEBUG"
}

function Save-Item {
    param($item, [string]$driveId, [string]$path, [string]$type, [bool]$isFolder, [long]$size)

    $fields = $null
    $listItemId = $null
    if ($CaptureItemFields -and $item.listItem) {
        $fields = $item.listItem.fields
        $listItemId = $item.listItem.id
    }
    $script:Counts.items++
    Write-Jsonl $script:Writers['items.jsonl'] ([ordered]@{
        library = $script:CurrentLibrary; path = $path; itemType = $type; isFolder = $isFolder; size = $size
        driveId = $driveId; itemId = $item.id; listItemId = $listItemId; fields = $fields
    })
}

function Get-Children-With-Paging {
    param(
        [string]$driveId,
        [string]$itemId,
        [string]$path = "/"
    )

    $baseUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$itemId/children"
    $useExpand = ($CaptureItemFields -and $script:ExpandFieldsSupported)
    $url = if ($useExpand) { $baseUrl + '?$expand=listItem($expand=fields)' } else { $baseUrl }
    Write-Log ("Lecture enfants: parent={0}" -f $path) "DEBUG"

    do {
        try {
            $response = Invoke-With-Retry -Uri $url
        } catch {
            if (-not $useExpand) { throw }
            $script:ExpandFieldsSupported = $false
            $useExpand = $false
            Write-Log ("Expansion listItem/fields refusee ({0}), metadonnees de colonnes desactivees" -f $_.Exception.Message) "WARN"
            $response = Invoke-With-Retry -Uri $baseUrl
        }
        if (-not $response) { return }

        foreach ($item in $response.value) {
            $itemPath = "$path$($item.name)"
            if ($item.folder) {
                $itemPath = "$itemPath/"
                [void]$script:allPaths.Add($itemPath)
                Save-Item -item $item -driveId $driveId -path $itemPath -type "Folder" -isFolder $true -size 0
                Get-Item-Permissions -driveId $driveId -itemId $item.id -path $itemPath -type "Folder"
                Get-Children-With-Paging -driveId $driveId -itemId $item.id -path $itemPath
            } else {
                [void]$script:allPaths.Add($itemPath)
                $size = 0
                try { $size = [long]$item.size } catch { $size = 0 }
                Save-Item -item $item -driveId $driveId -path $itemPath -type "File" -isFolder $false -size $size
                Get-Item-Permissions -driveId $driveId -itemId $item.id -path $itemPath -type "File"
            }
        }

        $url = $response.'@odata.nextLink'
    } while ($url)
}

function Scan-Library {
    param($list, [string]$key)

    $drive = Invoke-With-Retry -Uri "https://graph.microsoft.com/v1.0/sites/$($script:SiteId)/lists/$($list.id)/drive?`$select=id,name,webUrl"
    if (-not $drive -or -not $drive.id) { throw "Drive introuvable pour la bibliotheque '$key'" }

    $script:CurrentLibrary = $key
    # Chemins prefixes par la bibliotheque : "Shared Documents/dossier/fichier.docx".
    Get-Children-With-Paging -driveId $drive.id -itemId "root" -path "$key/"
    $script:CurrentLibrary = $null
    return $drive.id
}

function Scan-GenericList {
    param($list, [string]$key)

    $url = "https://graph.microsoft.com/v1.0/sites/$($script:SiteId)/lists/$($list.id)/items?`$expand=fields&`$top=500"
    $count = 0
    $writer = Open-Jsonl ("listitems\{0}.jsonl" -f $list.id)

    try {
        do {
            $resp = Invoke-With-Retry -Uri $url
            if (-not $resp) { break }
            foreach ($it in $resp.value) {
                Write-Jsonl $writer ([ordered]@{ id = [int]$it.id; fields = $it.fields })
                $count++
            }
            $url = $resp.'@odata.nextLink'
        } while ($url)
    } finally {
        $writer.Dispose()
        $script:Writers.Remove(("listitems\{0}.jsonl" -f $list.id))
    }

    $script:Counts.listItems += $count
    Write-Log ("Liste '{0}': {1} element(s)" -f $key, $count)
    return $count
}

function Export-MappingFile {
    if (-not $mappingPath) { return }
    if ($script:mappingRows.Count -eq 0) { Write-Log "Aucune entree mapping a exporter" "WARN"; return }

    # Conserve les destinations deja saisies dans un fichier existant.
    if (Test-Path -LiteralPath $mappingPath) {
        try {
            foreach ($old in (Import-Csv -LiteralPath $mappingPath -Encoding UTF8)) {
                if ($old.DestinationDisplayName -and $script:mappingRows.ContainsKey([string]$old.SourceDisplayName)) {
                    $script:mappingRows[[string]$old.SourceDisplayName].DestinationDisplayName = [string]$old.DestinationDisplayName
                }
            }
        } catch {
            Write-Log ("Ancien mapping illisible: {0}" -f $_.Exception.Message) "WARN"
        }
    }

    $script:mappingRows.Values | Sort-Object SourceDisplayName | Export-Csv -Path $mappingPath -NoTypeInformation -Encoding UTF8
    Write-Log "Mapping CSV genere (a completer: DestinationDisplayName): $mappingPath" "SUCCESS"
}

# =========================
# MAIN
# =========================
try {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    $script:LogFile = Join-Path $OutputDirectory 'discovery.log'
    Set-Content -LiteralPath $script:LogFile -Value '' -Encoding UTF8
    Write-Log "Demarrage scan SharePoint -> fichiers ($OutputDirectory)"

    $null = Open-Jsonl 'items.jsonl'
    $null = Open-Jsonl 'permissions.jsonl'
    $null = Open-Jsonl 'errors.jsonl'

    Write-Log "Connexion au site SharePoint"
    $site = Invoke-With-Retry -Uri $siteUrl
    if (-not $site) { throw "Connexion au site impossible" }

    $siteId = $site.id
    $script:SiteId = $siteId
    $script:SiteWebUrl = $site.webUrl

    $allLists = @(Get-AllPages -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists?`$select=id,name,displayName,description,webUrl,list")
    if ($allLists.Count -eq 0) { throw "Aucune liste/bibliotheque trouvee sur le site" }
    Write-Log ("Listes et bibliotheques trouvees: {0}" -f $allLists.Count)

    $listsOut = New-Object System.Collections.Generic.List[object]

    foreach ($list in $allLists) {
        $key = Get-ListRelativeKey -WebUrl $list.webUrl -SiteWebUrl $script:SiteWebUrl
        $hidden = $false
        try { $hidden = [bool]$list.list.hidden } catch {}

        if (-not (Test-ListSelected -Key $key -DisplayName ([string]$list.displayName) -Name ([string]$list.name) -Hidden $hidden)) {
            Write-Log ("Liste ignoree: {0}" -f $key) "DEBUG"
            continue
        }

        $template = [string]$list.list.template
        $isLibrary = ($template -match '(?i)library$')
        Write-Log ("Traitement {0}: {1} (template {2})" -f $(if ($isLibrary) { 'bibliotheque' } else { 'liste' }), $key, $template)

        try {
            $columns = @(Get-AllPages -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$($list.id)/columns")
            $driveId = $null
            $itemCount = 0

            if ($isLibrary) {
                $driveId = @(Scan-Library -list $list -key $key)[-1]
            } elseif ($ScanListItems) {
                $itemCount = [int]@(Scan-GenericList -list $list -key $key)[-1]
            }

            $listsOut.Add([ordered]@{
                listId      = [string]$list.id
                displayName = [string]$list.displayName
                relativeUrl = $key
                template    = $template
                isLibrary   = $isLibrary
                driveId     = $driveId
                description = [string]$list.description
                itemCount   = $itemCount
                columns     = $columns
            }) | Out-Null
        } catch {
            $script:CurrentLibrary = $null
            Write-Log ("Erreur sur '{0}': {1}" -f $key, $_.Exception.Message) "ERROR"
            Add-ErrorRow -path $key -phase "scan" -message $_.Exception.Message
        }
    }

    ConvertTo-Json -InputObject @($listsOut) -Depth 12 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'lists.json') -Encoding UTF8
    Write-Log ("lists.json ecrit: {0} liste(s)/bibliotheque(s)" -f $listsOut.Count) "SUCCESS"

    if ($enableCsvExport) {
        $script:permRowsForCsv | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
        Write-Log "Export CSV permissions: $exportPath" "SUCCESS"
    }

    Export-MappingFile

    Write-Log ("Scan termine: {0} fichiers/dossiers, {1} droits, {2} elements de liste, {3} erreur(s)" -f $script:Counts.items, $script:Counts.permissions, $script:Counts.listItems, $script:Counts.errors) "SUCCESS"
}
catch {
    $msg = $_.Exception.Message
    $inner = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { "" }
    Write-Log "Erreur critique: $msg" "ERROR"
    if ($inner) { Write-Log "Inner: $inner" "ERROR" }
    if ($_.ScriptStackTrace) { Write-Log "Stack: $($_.ScriptStackTrace)" "ERROR" }
    try { Export-MappingFile } catch {}
    exit 1
}
finally {
    Close-AllWriters
}
