#requires -Version 5.1

<#
.SYNOPSIS
    Migration SharePoint complete (source -> cible) en un seul script.

.DESCRIPTION
    Regroupe six phases, lancees dans cet ordre (choix via $Phases) :
      Decouverte   inventaire des listes, bibliotheques, fichiers, colonnes et droits (Graph) -> fichiers
      Copie        creation des listes/bibliotheques, copie des fichiers, elements, metadonnees et droits (Graph)
      Applications apps SharePoint installees sur le site (PnP)
      Navigation   menu de gauche : listes/bibliotheques et liens (PnP)
      Droits       niveaux d'autorisation, groupes SharePoint, droits uniques personnalises (PnP)
      Metadonnees  dates Cree/Modifie et auteurs Cree par/Modifie par d'origine (PnP), toujours en dernier

    Tous les parametres sont dans la section CONFIGURATION ci-dessous (aucun argument en ligne de commande).
    Les phases PnP simulent par defaut ($Apply = $false) ; mettre $Apply = $true pour appliquer.
    PnP.PowerShell 2.x/3.x exige PowerShell 7 : si une phase PnP est demandee depuis Windows
    PowerShell 5.1, le script se relance automatiquement dans pwsh.exe.

.EXAMPLE
    .\MigrationSharePoint.ps1
#>

# ============================================================================
# CONFIGURATION A MODIFIER
# ============================================================================

# Tous les parametres sont ici : modifiez ce bloc puis lancez simplement .\MigrationSharePoint.ps1
# --- Phases a executer. Ordre d'execution fixe : Decouverte, Copie, Applications, Navigation, Droits, Metadonnees.
#   Decouverte   : inventaire des listes/bibliotheques/fichiers/droits Graph -> fichiers (SharePoint-Discovery)
#   Copie        : creation des listes, copie des fichiers/elements/metadonnees + droits Graph (ecrit toujours)
#   Metadonnees  : remet Cree/Modifie (dates) et Cree par/Modifie par d'origine           (PnP, PowerShell 7)
#   Applications : apps SharePoint (SPFx) installees sur le site source -> cible       (PnP, PowerShell 7)
#   Navigation   : menu de gauche (listes/bibliotheques + liens)                        (PnP, PowerShell 7)
#   Droits       : niveaux d'autorisation, groupes, droits uniques personnalises        (PnP, PowerShell 7)
$Phases = @('Decouverte', 'Copie', 'Applications', 'Navigation', 'Droits', 'Metadonnees')

# $false = les phases PnP (Applications, Navigation, Droits) SIMULENT sans rien modifier sur la cible.
# La phase Copie n'a pas de simulation : elle ecrit toujours sur la cible.
$Apply = $false

# $false = on s'arrete a la premiere phase en echec ; $true = on enchaine quand meme.
$ContinueOnPhaseError = $false

# Dossier de travail (les sous-dossiers de chaque phase y sont crees).
$baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$WorkDirectory = $baseDir
$DiscoveryDirectory = Join-Path $WorkDirectory 'SharePoint-Discovery'

# --- Sites et connexions -----------------------------------------------------
# Graph (Decouverte, Copie) : secret client. PnP (Applications, Navigation, Droits) : certificat .pfx.
# ATTENTION : secrets en clair ; restreignez les droits NTFS du script. Ne le publiez pas renseigne.
$SourceSiteUrl = 'https://infoprodigital365.sharepoint.com/sites/fethygate'
$TargetSiteUrl = 'https://infoprodigital365.sharepoint.com/sites/fethiMercury'

$SourceTenantId = '2eea08b8-1972-447b-ad43-d044d042500a'
$SourceClientId = '821109a4-6e0e-48e8-b477-8f9b70aa32a4'
$SourceClientSecret = ''
$SourceCertificatePath = 'C:\certs\MSGraphExchangeOnlineAuth20260717_2.pfx'
$SourceCertificatePassword = ''

# Vide = memes valeurs que la source (copie dans le meme tenant).
$TargetTenantId = ''
$TargetClientId = ''
$TargetClientSecret = ''
$TargetCertificatePath = ''
$TargetCertificatePassword = ''

# --- Perimetre commun : listes et bibliotheques ------------------------------
# Les listes techniques/cachees ne sont pas traitees par defaut (Decouverte et Droits).
$IncludeHiddenLists = $false

# --- Phase Decouverte ----------------------------------------------------------
# Vide = tout le site. Sinon noms (titre ou segment d'URL, ex. "Documents", "Shared Documents", "Lists/Taches").
$ListNameFilter = @()
# Chemins relatifs (jokers autorises) toujours ignores.
$ExcludedListUrlPatterns = @('appdata', 'fpdatasources', 'FormServerTemplates', 'Style Library', '_catalogs/*', 'SiteCollectionImages', 'Lists/DO_NOT_DELETE*')
# $true = enregistre aussi les elements des listes classiques.
$ScanListItems = $true
# $true = enregistre les valeurs des colonnes (metadonnees) des fichiers/dossiers.
$CaptureItemFields = $true
# Exports CSV (le mapping est a completer : colonne DestinationDisplayName, puis relu par la Copie).
$enableCsvExport = $true
$exportPath  = "E:\SharePoint_Permissions_Export_src9999.csv"
$mappingPath = "E:\UserMapping.csv"

# --- Phase Copie -----------------------------------------------------------------
# $false (defaut) = aucune traduction de chemin, les chemins sont copies a l'identique.
# $true = applique les traductions de $translation_csv (colonnes Source, Destination).
$ApplyPathTranslations = $false
$translation_csv = "E:\UKREiiF_FinalCopy_path.csv"
$CsvLibraryPrefixToStrip = ""
$ForceOverwrite = $true
$ChunkSize      = 8MB
# $true = cree sur la cible les listes/bibliotheques absentes (meme URL relative, colonnes personnalisees).
$CreateMissingLists = $true
# $true = copie les elements des listes classiques (les ID sont preserves au mieux, voir phase Droits).
$CopyListItems = $true
# $true = copie les valeurs des colonnes personnalisees des fichiers/dossiers et des elements.
$CopyItemFields = $true
# Nombre max d'elements factices crees pour conserver les ID d'une liste (trous d'ID dans la source).
$MaxIdPlaceholders = 5000
# Renommage eventuel d'une liste/bibliotheque : @{ "Shared Documents" = "Documents partages" } (cle = URL relative source).
# Attention : la phase Droits retrouve les objets par URL relative identique.
$LibraryNameMapping = @{}
# Colonnes jamais copiees (systeme / calculees par SharePoint).
$FieldSkipNames = @('ID', 'Id', 'ContentType', 'ContentTypeId', 'Attachments', 'FileLeafRef', 'FileRef', 'FileDirRef', 'LinkFilename', 'LinkFilenameNoMenu',
    'Created', 'Modified', 'Author', 'Editor', 'AppAuthor', 'AppEditor', '_UIVersionString', 'Edit', 'DocIcon', 'FolderChildCount', 'ItemChildCount',
    'LinkTitle', 'LinkTitleNoMenu', 'ComplianceAssetId', '_ColorTag', 'ParentLeafName', 'ParentVersionString')
# Groupes de colonnes natifs (deja presents dans la liste creee depuis le meme modele).
$NativeColumnGroups = @('_Hidden', 'Core Document Columns', 'Core Contact and Calendar Columns', 'Core Task and Issue Columns', 'Base Columns',
    'Document and Record Management Columns', 'Extended Columns', 'Display Template Columns', 'Enterprise Keywords Group', 'Publishing Columns')

# --- Phase Applications ------------------------------------------------------------
# Catalogues a examiner : 'Tenant' et/ou 'Site' (catalogue de la collection de sites).
$Scopes = @('Tenant', 'Site')
# Dossier de packages .sppkg a deployer sur la cible quand l'app n'y est pas disponible (optionnel).
$SppkgFolder = ''

# --- Phase Navigation ----------------------------------------------------------------
# Emplacement a copier : QuickLaunch (menu de gauche) ou TopNavigationBar (barre du haut).
$Locations = @('QuickLaunch')

# --- Phase Metadonnees ------------------------------------------------------------------
# Remet sur la cible les valeurs d'origine que la copie ne peut pas ecrire (Graph) :
# Cree / Modifie (dates) et Cree par / Modifie par. Necessite d'avoir relance Decouverte + Copie
# avec cette version. Les adresses e-mail sont transposees avec $DomainMapping (section Droits ci-dessous).
$MetadataSetDates = $true
$MetadataSetAuthors = $true
# $true = relit chaque element apres ecriture et signale (log + rapport) les valeurs que SharePoint n'a pas retenues.
$MetadataVerify = $true

# --- Phase Droits (autorisations non natives) ------------------------------------------
# $true = remplace les droits uniques des listes, dossiers, fichiers et elements.
# Les droits racine du site restent fusionnes pour eviter un verrouillage.
$ReplaceUniquePermissions = $false

# Mettre $true pour ne pas analyser les droits individuels des dossiers/fichiers/elements.
$SkipItemPermissions = $false

# Correspondance de domaines pour les comptes internes du tenant source.
# Exemple : @{ 'source.com' = 'cible.com'; 'source.onmicrosoft.com' = 'cible.com' }
# L'adresse transformee est cherchee en premier, puis l'adresse d'origine.
$DomainMapping = @{}

# $true = cree en invite B2B (sans e-mail d'invitation) les comptes introuvables
# dans le tenant cible. $false = ils sont listes dans UnresolvedPrincipals.csv.
$InviteMissingUsers = $false

# $true = les droits donnes aux groupes natifs Proprietaires/Membres/Visiteurs sur
# des objets a droits uniques (ou avec un niveau personnalise) sont reportes sur
# les groupes natifs equivalents du site cible. Leurs membres ne sont PAS copies.
$IncludeAssociatedGroupAssignments = $true

# $true (MIROIR) = pour chaque objet a droits uniques dans la source, les droits de
# l'objet cible sont REMPLACES par ceux de la source (qu'il herite ou qu'il ait deja
# des droits uniques, par exemple copies par l'outil de migration du contenu).
# Groupes natifs et groupe Microsoft 365 du site sont transposes vers ceux du site cible.
# $false = les droits source sont AJOUTES aux droits existants de la cible (fusion).
$MirrorSourcePermissions = $true

# $true = reporte les droits donnes a "Tout le monde" / "Tout le monde sauf les
# utilisateurs externes" (necessite un GUID dans $TargetTenantId pour le second).
$MigrateEveryoneClaims = $false

# PARTAGES VERS DES ADRESSES E-MAIL EXTERNES (aucun e-mail n'est envoye) :
# les destinataires des liens de partage "Personnes specifiques" du site source
# recoivent sur l'objet cible un droit DIRECT equivalent (Lecture / Modification...).
# Les liens eux-memes (URL) ne peuvent pas etre recrees a l'identique.
# 'External' = uniquement les destinataires externes (invites, adresses hors tenant) ;
# 'All'      = tous les destinataires (internes et externes) ;
# 'None'     = les liens de partage sont ignores.
$SharingLinkRecipients = 'External'

# $true = un destinataire externe absent du tenant cible est cree comme invite B2B
# SANS e-mail d'invitation (sendInvitationMessage = false), meme si
# $InviteMissingUsers = $false. Necessite User.Invite.All (Graph) sur l'application cible.
$InviteExternalRecipients = $true

# Groupes SharePoint consideres comme natifs/systeme : jamais recrees.
$ExcludedGroupTitlePatterns = @(
    'SharingLinks.*'
    'Limited Access System Group*'
    'Excel Services Viewers', 'Visionneuses Excel Services'
    'Approvers', 'Approbateurs'
    'Designers', 'Concepteurs'
    'Hierarchy Managers', 'Gestionnaires de hi*rarchie'
    'Restricted Readers', 'Lecteurs restreints'
    'Style Resource Readers', 'Lecteurs des ressources de style'
    'Translation Managers', 'Gestionnaires de traduction'
    'Quick Deploy Users', 'Utilisateurs de d*ploiement rapide'
)

# Niveaux natifs dont RoleTypeKind vaut None (non detectables autrement).
# Chaque ligne regroupe les noms equivalents selon la langue du site.
$NativeRoleNameGroups = @(
    @('View Only', 'Affichage seul'),
    @('Approve', 'Approuver'),
    # [char]0xE9 = e accent aigu (evite les problemes d'encodage sous ISE / PowerShell 5.1).
    @('Manage Hierarchy', "G$([char]0xE9)rer la hi$([char]0xE9)rarchie"),
    @('Restricted Read', 'Lecture restreinte'),
    @('Restricted Interfaces for Translation', 'Interfaces restreintes pour la traduction'),
    @('Restricted View', 'Affichage restreint')
)

# Laisser vide tant qu'aucun fichier de correspondance n'est necessaire.
# Format CSV attendu : SourceKey,TargetLogin (voir UnresolvedPrincipals.csv pour les SourceKey).
# Pour une cle SPGROUP:..., TargetLogin est le titre du groupe SharePoint cible.
$PrincipalMappingCsv = ''

# Niveau de detail des logs a l'ecran et dans Suivi_<date>.log :
# 'Normal' = etapes, actions, avertissements et erreurs ;
# 'Detail' = en plus : chaque liste, element, principal et recherche Graph.
$LogLevel = 'Detail'

# ============================================================================
# FIN DE LA CONFIGURATION - NE PAS MODIFIER LA SUITE
# ============================================================================

# ============================================================================
# PREPARATION
# ============================================================================
if ([string]::IsNullOrWhiteSpace($TargetTenantId)) { $TargetTenantId = $SourceTenantId }
if ([string]::IsNullOrWhiteSpace($TargetClientId)) { $TargetClientId = $SourceClientId }
if ([string]::IsNullOrWhiteSpace($TargetClientSecret)) { $TargetClientSecret = $SourceClientSecret }
if ([string]::IsNullOrWhiteSpace($TargetCertificatePath)) { $TargetCertificatePath = $SourceCertificatePath }
if ([string]::IsNullOrWhiteSpace($TargetCertificatePassword)) { $TargetCertificatePassword = $SourceCertificatePassword }

$AllPhases = @('Decouverte', 'Copie', 'Applications', 'Navigation', 'Droits', 'Metadonnees')
foreach ($p in $Phases) {
    if ($AllPhases -notcontains $p) { throw "Phase inconnue '$p'. Valeurs possibles : $($AllPhases -join ', ')" }
}
$PnpPhases = @('Metadonnees', 'Applications', 'Navigation', 'Droits')
$needsPnp = @($Phases | Where-Object { $PnpPhases -contains $_ }).Count -gt 0

# PnP.PowerShell 2.x/3.x exige PowerShell 7 : relance automatique dans pwsh.exe.
if ($needsPnp -and $PSVersionTable.PSVersion.Major -lt 7) {
    $pwshCommand = Get-Command -Name pwsh.exe -ErrorAction SilentlyContinue
    if ($pwshCommand -and $PSCommandPath) {
        Write-Host "Windows PowerShell $($PSVersionTable.PSVersion) detecte : relance dans PowerShell 7 ($($pwshCommand.Source))..." -ForegroundColor Cyan
        & $pwshCommand.Source -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath
        exit $LASTEXITCODE
    }
}

$pnpModule = $null
if ($needsPnp) {
    $pnpModule = Get-Module -ListAvailable -Name PnP.PowerShell | Sort-Object Version -Descending | Select-Object -First 1
    if ($null -eq $pnpModule) {
        throw 'Module PnP.PowerShell absent. Installez PowerShell 7 puis, dans pwsh : Install-Module PnP.PowerShell -Scope CurrentUser'
    }
    if ($PSVersionTable.PSVersion.Major -lt 7 -and $pnpModule.Version.Major -ge 2) {
        throw 'PnP.PowerShell 2.x/3.x exige PowerShell 7 : installez PowerShell 7 (winget install --id Microsoft.PowerShell) ou PnP.PowerShell 1.12.0.'
    }
    Import-Module -Name PnP.PowerShell -RequiredVersion $pnpModule.Version -DisableNameChecking
}

# ============================================================================
# PHASES
# ============================================================================

function Invoke-PhaseDecouverte {
    # Phase Decouverte : inventaire Graph -> fichiers
    try {
        if ($Host -and $Host.Name -and $Host.Name -notmatch 'ISE') {
            if (-not [Console]::IsInputRedirected) { [Console]::InputEncoding = [System.Text.Encoding]::UTF8 }
            if (-not [Console]::IsOutputRedirected) { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 }
        }
    } catch {}
    $OutputDirectory = Join-Path $WorkDirectory 'SharePoint-Discovery'
    $tenantId     = $SourceTenantId
    $clientId     = $SourceClientId
    $clientSecret = $SourceClientSecret
    $hostAndPath  = [Uri]$SourceSiteUrl
    $siteUrl = "https://graph.microsoft.com/v1.0/sites/$($hostAndPath.Host)/:$($hostAndPath.AbsolutePath.TrimEnd('/'))"

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
            try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -ErrorAction Stop } catch {}
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

    # Auteur/date de creation et de derniere modification d'un element Graph (relus par la phase Metadonnees).
    function Convert-IsoDate($value) {
        if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) { return $null }
        try { return ([datetime]$value).ToUniversalTime().ToString('o') } catch { return [string]$value }
    }
    function Get-Identity($identitySet, $fallbackSet) {
        $u = $null
        if ($identitySet -and $identitySet.user) { $u = $identitySet.user }
        if ((-not $u -or -not $u.email) -and $fallbackSet -and $fallbackSet.user) { $u = $fallbackSet.user }
        if (-not $u) { return $null }
        return [ordered]@{ email = [string]$u.email; name = [string]$u.displayName; id = [string]$u.id }
    }
    function Get-ItemMetadata($item) {
        $li = $item.listItem
        return [ordered]@{
            created    = Convert-IsoDate $item.createdDateTime
            modified   = Convert-IsoDate $item.lastModifiedDateTime
            createdBy  = Get-Identity $item.createdBy $(if ($li) { $li.createdBy })
            modifiedBy = Get-Identity $item.lastModifiedBy $(if ($li) { $li.lastModifiedBy })
        }
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
            meta = (Get-ItemMetadata $item)
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
                    Write-Jsonl $writer ([ordered]@{ id = [int]$it.id; fields = $it.fields; meta = (Get-ItemMetadata $it) })
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

                $listsOut.Add([pscustomobject][ordered]@{
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

        # Serialisation element par element puis assemblage du tableau (evite les erreurs de ConvertTo-Json sur de gros tableaux).
        $listJsonParts = New-Object System.Collections.Generic.List[string]
        foreach ($l in $listsOut) { $listJsonParts.Add((ConvertTo-Json -InputObject $l -Depth 12 -Compress)) }
        $listsJson = '[' + ($listJsonParts -join ',') + ']'
        [System.IO.File]::WriteAllText((Join-Path $OutputDirectory 'lists.json'), $listsJson, (New-Object System.Text.UTF8Encoding($false)))
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
        $script:PhaseFailed = $true
    }
    finally {
        Close-AllWriters
    }
}

function Invoke-PhaseCopie {
    # Phase Copie : listes, fichiers, elements, metadonnees et droits Graph
    $connexion_source = @{ tenant_id = $SourceTenantId; client_id = $SourceClientId; client_secret = $SourceClientSecret }
    $connexion_cible  = @{ tenant_id = $TargetTenantId; client_id = $TargetClientId; client_secret = $TargetClientSecret }
    $site_url_source = $SourceSiteUrl
    $site_url_cible  = $TargetSiteUrl
    $mapping_csv     = $mappingPath
    $OutputDirectory = Join-Path $WorkDirectory 'SharePoint-Copy'

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
            try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -ErrorAction Stop } catch {}
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
        if (-not $ApplyPathTranslations) {
            Write-Log "Traduction de chemins desactivee (ApplyPathTranslations = false)"
            return
        }
        if (-not (Test-Path $translation_csv)) {
            Write-Log "Fichier de traduction introuvable: $translation_csv" "WARN"
            return
        }

        $csv = Import-Csv -Path $translation_csv -Encoding UTF8
        foreach ($r in $csv) {
            if ($r.Source -and $r.Destination) {
                $src = Normalize-RelPath $r.Source
                if ([string]::IsNullOrEmpty($src)) {
                    # Une source vide (ex. "/") remplacerait n'importe quel chemin : ignoree.
                    Write-Log ("Traduction ignoree (source vide apres normalisation): '{0}' -> '{1}'" -f $r.Source, $r.Destination) "WARN"
                    continue
                }
                $script:UrlTranslations[$src] = (Normalize-RelPath $r.Destination)
            }
        }
        Write-Log ("Traductions chargees: {0}" -f $script:UrlTranslations.Count) "SUCCESS"
    }

    function Apply-UrlTranslations {
        param([string]$Path)

        $p = Normalize-RelPath $Path
        if (-not $ApplyPathTranslations) { return $p }
        foreach ($k in $script:UrlTranslations.Keys) {
            if (-not [string]::IsNullOrEmpty($k) -and $p.Contains($k)) {
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
            [object[]]$permRows = @()
            if ($permsByPath.ContainsKey($path)) { $permRows = $permsByPath[$path].ToArray() }

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
        $script:PhaseFailed = $true
    }
    finally {
        try { Save-Results } catch { Write-Host "Echec ecriture des resultats: $($_.Exception.Message)" }
        if ($script:HttpClient) { $script:HttpClient.Dispose() }
    }
}

function Invoke-PhaseApplications {
    # Phase Applications : apps SharePoint (SPFx) du site source -> cible
    $OutputDirectory = Join-Path $WorkDirectory 'SharePoint-Apps'
    $ErrorActionPreference = 'Stop'

    $script:Results = New-Object System.Collections.Generic.List[object]
    $script:LogFile = $null

    function Write-Log {
        param([string]$Message, [string]$Level = 'INFO')
        $line = "[{0}][{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
        $color = switch ($Level) { 'ERROR' { 'Red' } 'WARN' { 'Yellow' } 'SUCCESS' { 'Green' } default { 'White' } }
        Write-Host $line -ForegroundColor $color
        if ($script:LogFile) { try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -ErrorAction Stop } catch {} }
    }

    function Add-Result {
        param([string]$Scope, $App, [string]$Status, [string]$Detail)
        $script:Results.Add([pscustomobject]@{
            Scope = $Scope; AppId = [string]$App.Id; Title = [string]$App.Title
            SourceVersion = [string]$App.InstalledVersion; Status = $Status; Detail = $Detail
        }) | Out-Null
    }

    function Test-Installed($App) {
        $v = [string]$App.InstalledVersion
        return (-not [string]::IsNullOrWhiteSpace($v) -and $v -ne '0.0.0.0')
    }

    function Connect-Site {
        param([string]$Url, [string]$Tenant, [string]$ClientId, [string]$CertPath, [string]$CertPassword)
        if (-not (Test-Path -LiteralPath $CertPath -PathType Leaf)) { throw "Certificat introuvable: $CertPath" }
        if ([string]::IsNullOrWhiteSpace($CertPassword)) { throw "Mot de passe du certificat vide pour $CertPath" }
        return Connect-PnPOnline -Url $Url -Tenant $Tenant -ClientId $ClientId -CertificatePath $CertPath `
            -CertificatePassword (ConvertTo-SecureString -String $CertPassword -AsPlainText -Force) -ReturnConnection
    }

    function Get-CatalogApps {
        param([string]$Scope, $Connection)
        try {
            return @(Get-PnPApp -Scope $Scope -Connection $Connection)
        } catch {
            Write-Log ("Catalogue '{0}' illisible: {1}" -f $Scope, $_.Exception.Message) 'WARN'
            return @()
        }
    }

    try {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        $script:LogFile = Join-Path $OutputDirectory 'applications.log'
        Set-Content -LiteralPath $script:LogFile -Value '' -Encoding UTF8
        Write-Log ("Mode: {0}" -f $(if ($Apply) { 'APPLICATION' } else { 'SIMULATION (aucune modification)' }))

        if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
            throw 'Module PnP.PowerShell absent. Installez PowerShell 7 puis, dans pwsh : Install-Module PnP.PowerShell -Scope CurrentUser'
        }
        Import-Module PnP.PowerShell -DisableNameChecking

        Write-Log "Connexion au site source: $SourceSiteUrl"
        $src = Connect-Site -Url $SourceSiteUrl -Tenant $SourceTenantId -ClientId $SourceClientId -CertPath $SourceCertificatePath -CertPassword $SourceCertificatePassword
        Write-Log "Connexion au site cible: $TargetSiteUrl"
        $dst = Connect-Site -Url $TargetSiteUrl -Tenant $TargetTenantId -ClientId $TargetClientId -CertPath $TargetCertificatePath -CertPassword $TargetCertificatePassword

        # ---- Source : apps installees ----
        $sourceApps = New-Object System.Collections.Generic.List[object]
        foreach ($scope in $Scopes) {
            $apps = Get-CatalogApps -Scope $scope -Connection $src
            $installed = @($apps | Where-Object { Test-Installed $_ })
            Write-Log ("Source / catalogue {0}: {1} app(s), dont {2} installee(s) sur le site" -f $scope, $apps.Count, $installed.Count)
            foreach ($a in $installed) { $sourceApps.Add([pscustomobject]@{ Scope = $scope; App = $a }) | Out-Null }
        }

        if ($sourceApps.Count -eq 0) {
            Write-Log 'Aucune app installee sur le site source (hors add-ins classiques).' 'WARN'
        }

        # ---- Cible : catalogue + installation ----
        $targetCatalog = @{}
        foreach ($scope in $Scopes) {
            $targetCatalog[$scope] = Get-CatalogApps -Scope $scope -Connection $dst
        }

        foreach ($entry in $sourceApps) {
            $scope = $entry.Scope
            $app = $entry.App
            $label = "{0} [{1}]" -f $app.Title, $scope
            try {
                $found = @($targetCatalog[$scope] | Where-Object { [string]$_.Id -eq [string]$app.Id }) | Select-Object -First 1

                if (-not $found -and -not [string]::IsNullOrWhiteSpace($SppkgFolder)) {
                    $pkg = Get-ChildItem -LiteralPath $SppkgFolder -Filter '*.sppkg' -File -ErrorAction SilentlyContinue |
                        Where-Object { $_.BaseName -ieq [string]$app.Title -or $_.BaseName -ieq [string]$app.Id } | Select-Object -First 1
                    if ($pkg) {
                        if ($Apply) {
                            $null = Add-PnPApp -Path $pkg.FullName -Scope $scope -Publish -Connection $dst
                            Write-Log ("Package deploye: {0} ({1})" -f $pkg.Name, $scope) 'SUCCESS'
                            $targetCatalog[$scope] = Get-CatalogApps -Scope $scope -Connection $dst
                            $found = @($targetCatalog[$scope] | Where-Object { [string]$_.Id -eq [string]$app.Id }) | Select-Object -First 1
                        } else {
                            Add-Result -Scope $scope -App $app -Status 'Planned' -Detail "Deploiement de $($pkg.Name) puis installation"
                            Write-Log ("[Simulation] deploiement de {0} puis installation" -f $pkg.Name)
                            continue
                        }
                    }
                }

                if (-not $found) {
                    Add-Result -Scope $scope -App $app -Status 'NonDisponible' -Detail 'App absente du catalogue cible : deployer le .sppkg (voir $SppkgFolder)'
                    Write-Log ("{0} : absente du catalogue cible" -f $label) 'WARN'
                    continue
                }

                if (Test-Installed $found) {
                    $note = if ($found.CanUpgrade) { 'deja installee (mise a jour disponible)' } else { 'deja installee' }
                    Add-Result -Scope $scope -App $found -Status 'DejaInstallee' -Detail $note
                    Write-Log ("{0} : {1}" -f $label, $note)
                    continue
                }

                if (-not $Apply) {
                    Add-Result -Scope $scope -App $app -Status 'Planned' -Detail 'Installation prevue'
                    Write-Log ("[Simulation] installation de {0}" -f $label)
                    continue
                }

                Install-PnPApp -Identity ([guid]$app.Id) -Scope $scope -Wait -Connection $dst
                Add-Result -Scope $scope -App $app -Status 'Installee' -Detail ''
                Write-Log ("{0} : installee" -f $label) 'SUCCESS'
            } catch {
                Add-Result -Scope $scope -App $app -Status 'Erreur' -Detail $_.Exception.Message
                Write-Log ("{0} : {1}" -f $label, $_.Exception.Message) 'ERROR'
            }
        }
    }
    catch {
        Write-Log ("Erreur fatale: {0}" -f $_.Exception.Message) 'ERROR'
        $script:PhaseFailed = $true
    }
    finally {
        if ($script:Results.Count -gt 0 -and (Test-Path -LiteralPath $OutputDirectory)) {
            $csv = Join-Path $OutputDirectory 'applications_resultats.csv'
            $script:Results | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
            Write-Log ("Rapport: {0} ({1} ligne(s))" -f $csv, $script:Results.Count)
        }
    }
}

function Invoke-PhaseNavigation {
    # Phase Navigation : menu de gauche (listes/bibliotheques + liens)
    $ErrorActionPreference = 'Stop'

    function Write-Log {
        param([string]$Message, [string]$Level = 'INFO')
        $color = switch ($Level) { 'ERROR' { 'Red' } 'WARN' { 'Yellow' } 'SUCCESS' { 'Green' } default { 'White' } }
        Write-Host ("[{0}][{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message) -ForegroundColor $color
    }

    function Connect-Site {
        param([string]$Url, [string]$Tenant, [string]$ClientId, [string]$CertPath, [string]$CertPassword)
        if (-not (Test-Path -LiteralPath $CertPath -PathType Leaf)) { throw "Certificat introuvable: $CertPath" }
        if ([string]::IsNullOrWhiteSpace($CertPassword)) { throw "Mot de passe du certificat vide pour $CertPath" }
        return Connect-PnPOnline -Url $Url -Tenant $Tenant -ClientId $ClientId -CertificatePath $CertPath `
            -CertificatePassword (ConvertTo-SecureString -String $CertPassword -AsPlainText -Force) -ReturnConnection
    }

    function Convert-NodeUrl {
        param([string]$Url, [string]$SourceWebUrl, [string]$TargetWebUrl)
        if ([string]::IsNullOrWhiteSpace($Url)) { return $Url }
        if ($Url.StartsWith($SourceWebUrl, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $TargetWebUrl.TrimEnd('/') + $Url.Substring($SourceWebUrl.Length)
        }
        return $Url
    }

    function Add-MissingNodes {
        param($Nodes, $ParentId, [string]$Location, $Conn, [string]$SrcWeb, [string]$DstWeb, $ExistingByTitle)

        foreach ($node in @($Nodes)) {
            $title = [string]$node.Title
            $url = Convert-NodeUrl -Url ([string]$node.Url) -SourceWebUrl $SrcWeb -TargetWebUrl $DstWeb
            if ([string]::IsNullOrWhiteSpace($title) -or [string]::IsNullOrWhiteSpace($url)) {
                Write-Log ("noeud ignore (titre ou URL vide): '{0}' / '{1}'" -f $title, $node.Url) 'WARN'
                continue
            }

            $key = "$ParentId|$title".ToLowerInvariant()
            $newId = $null
            if ($ExistingByTitle.ContainsKey($key)) {
                $newId = $ExistingByTitle[$key]
                Write-Log ("deja presente: {0}" -f $title)
            } elseif (-not $Apply) {
                Write-Log ("[Simulation] ajout: {0} -> {1}" -f $title, $url)
            } else {
                $params = @{ Location = $Location; Title = $title; Url = $url; Connection = $Conn }
                if ($ParentId) { $params['Parent'] = [int]$ParentId }
                if ($url -match '^https?://' -and $url -notlike ($DstWeb.TrimEnd('/') + '*')) { $params['External'] = $true }
                $created = Add-PnPNavigationNode @params
                $newId = $created.Id
                Write-Log ("ajoutee: {0} -> {1}" -f $title, $url) 'SUCCESS'
            }

            $children = @($node.Children)
            if ($children.Count -gt 0 -and $newId) {
                Add-MissingNodes -Nodes $children -ParentId $newId -Location $Location -Conn $Conn -SrcWeb $SrcWeb -DstWeb $DstWeb -ExistingByTitle $ExistingByTitle
            }
        }
    }

    # Les bibliotheques/listes apparaissent dans le menu de gauche grace a leur propriete OnQuickLaunch
    # ("Afficher dans la navigation"), pas via des noeuds de navigation : on recopie cette propriete.
    function Sync-ListNavigation {
        param($Src, $Dst, [string]$SrcWeb, [string]$DstWeb)

        $props = 'Title', 'Hidden', 'IsSystemList', 'OnQuickLaunch', 'RootFolder'
        $srcLists = @(Get-PnPList -Includes $props -Connection $Src)
        $dstMap = @{}
        foreach ($l in @(Get-PnPList -Includes $props -Connection $Dst)) {
            $rel = ([string]$l.RootFolder.ServerRelativeUrl).Substring($DstWeb.Length).TrimStart('/').ToLowerInvariant()
            $dstMap[$rel] = $l
        }

        foreach ($l in $srcLists) {
            if ($l.Hidden -or $l.IsSystemList -or -not $l.OnQuickLaunch) { continue }
            $rel = ([string]$l.RootFolder.ServerRelativeUrl).Substring($SrcWeb.Length).TrimStart('/')
            $target = $dstMap[$rel.ToLowerInvariant()]
            if (-not $target) {
                Write-Log ("liste absente de la cible: {0}" -f $rel) 'WARN'
                continue
            }
            if ($target.OnQuickLaunch) {
                Write-Log ("deja dans la navigation: {0}" -f $rel)
            } elseif (-not $Apply) {
                Write-Log ("[Simulation] affichage dans la navigation: {0}" -f $rel)
            } else {
                # Set-PnPList n'expose pas ce parametre : on passe par l'objet CSOM.
                $target.OnQuickLaunch = $true
                $target.Update()
                Invoke-PnPQuery -Connection $Dst
                Write-Log ("ajoutee a la navigation: {0}" -f $rel) 'SUCCESS'
            }
        }
    }

    try {
        Write-Log ("Mode: {0}" -f $(if ($Apply) { 'APPLICATION' } else { 'SIMULATION (aucune modification)' }))
        if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) { throw 'Module PnP.PowerShell absent. Installez PowerShell 7 puis, dans pwsh : Install-Module PnP.PowerShell -Scope CurrentUser' }
        Import-Module PnP.PowerShell -DisableNameChecking

        $src = Connect-Site -Url $SourceSiteUrl -Tenant $SourceTenantId -ClientId $SourceClientId -CertPath $SourceCertificatePath -CertPassword $SourceCertificatePassword
        $dst = Connect-Site -Url $TargetSiteUrl -Tenant $TargetTenantId -ClientId $TargetClientId -CertPath $TargetCertificatePath -CertPassword $TargetCertificatePassword

        $srcWeb = (Get-PnPWeb -Connection $src).ServerRelativeUrl
        $dstWeb = (Get-PnPWeb -Connection $dst).ServerRelativeUrl

        Write-Log 'Listes et bibliotheques affichees dans la navigation'
        Sync-ListNavigation -Src $src -Dst $dst -SrcWeb $srcWeb -DstWeb $dstWeb

        foreach ($location in $Locations) {
            Write-Log ("Navigation {0}" -f $location)
            $srcNodes = @(Get-PnPNavigationNode -Location $location -Tree -Connection $src)
            Write-Log ("Source: {0} noeud(s) de premier niveau" -f $srcNodes.Count)

            # Index des noeuds existants sur la cible (parent|titre -> id).
            $existing = @{}
            foreach ($n in @(Get-PnPNavigationNode -Location $location -Tree -Connection $dst)) {
                $existing["|$($n.Title)".ToLowerInvariant()] = $n.Id
                foreach ($c in @($n.Children)) { $existing["$($n.Id)|$($c.Title)".ToLowerInvariant()] = $c.Id }
            }

            Add-MissingNodes -Nodes $srcNodes -ParentId $null -Location $location -Conn $dst -SrcWeb $srcWeb -DstWeb $dstWeb -ExistingByTitle $existing
        }
        Write-Log 'Termine' 'SUCCESS'
    }
    catch {
        Write-Log ("Erreur: {0}" -f $_.Exception.Message) 'ERROR'
        $script:PhaseFailed = $true
    }
}

function Invoke-PhaseDroits {
    # Phase Droits : niveaux d'autorisation, groupes, droits uniques personnalises (PnP)
    $OutputDirectory = Join-Path $WorkDirectory 'SharePoint-Permissions-Migration'

    # Parametres disponibles selon la version du module.
    $script:GraphSupportsEventual = (Get-Command Invoke-PnPGraphMethod).Parameters.ContainsKey('ConsistencyLevelEventual')
    $script:ItemPermissionSupportsSystemUpdate = (Get-Command Set-PnPListItemPermission).Parameters.ContainsKey('SystemUpdate')

    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'Continue'

    $script:MigrationLog = [System.Collections.Generic.List[object]]::new()
    $script:Unresolved = [System.Collections.Generic.List[object]]::new()
    $script:Invitations = [System.Collections.Generic.List[object]]::new()
    $script:PrincipalMappings = @{}
    $script:TargetPrincipalCache = @{}
    $script:TargetRoleCache = @{}
    $script:TargetGroupMap = @{}
    $script:SourceAssociatedGroups = @{}
    $script:SourceEmailCache = @{}
    $script:SkippedPrincipalKeys = @{}
    $script:TargetTenantGuid = $null
    $script:SameTenant = $false
    $script:SourceSiteGroupId = $null
    $script:TargetSiteGroupId = $null
    $script:SourceGraphSiteId = $null
    $script:ExternalEmailCache = @{}
    $script:StartTime = Get-Date
    $script:LiveLogFile = $null
    $script:StepStart = $null
    $script:StepName = $null
    $script:StepIndex = 0
    $script:StepCount = 5

    function Write-Log {
        param(
            [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
            [ValidateSet('Info', 'Detail', 'Success', 'Warn', 'Error', 'Step')][string]$Level = 'Info'
        )

        if ($Level -eq 'Detail' -and $LogLevel -ne 'Detail') {
            return
        }

        $elapsed = (Get-Date) - $script:StartTime
        $prefix = '{0:HH:mm:ss} [+{1:hh\:mm\:ss}] {2,-7}' -f (Get-Date), $elapsed, $Level.ToUpperInvariant()
        $line = "$prefix $Message"

        $color = switch ($Level) {
            'Error'   { 'Red' }
            'Warn'    { 'Yellow' }
            'Success' { 'Green' }
            'Step'    { 'Magenta' }
            'Detail'  { 'DarkGray' }
            default   { 'White' }
        }
        Write-Host $line -ForegroundColor $color

        # Ecriture immediate : le fichier reste exploitable meme si le script est interrompu.
        if ($null -ne $script:LiveLogFile) {
            try {
                Add-Content -LiteralPath $script:LiveLogFile -Value $line -Encoding utf8
            }
            catch {
                # Ne jamais interrompre la migration a cause du journal.
            }
        }
    }

    function Show-Progress {
        param(
            [Parameter(Mandatory)][string]$Activity,
            [AllowEmptyString()][AllowNull()][string]$Status,
            [int]$Index,
            [int]$Total
        )

        # La barre de progression ne doit jamais interrompre la migration.
        try {
            if ([string]::IsNullOrWhiteSpace($Status)) {
                $Status = '(sans nom)'
            }
            $percent = [Math]::Min(100, [int](($Index / [Math]::Max(1, $Total)) * 100))
            Write-Progress -Activity $Activity -Status "[$Index/$Total] $Status" -PercentComplete $percent
        }
        catch {
            # Ignore.
        }
    }

    function Start-Step {
        param([Parameter(Mandatory)][string]$Name)

        Complete-Step
        $script:StepIndex++
        $script:StepName = $Name
        $script:StepStart = Get-Date
        Write-Log -Level Step -Message ''
        Write-Log -Level Step -Message ("=" * 70)
        Write-Log -Level Step -Message "ETAPE $($script:StepIndex)/$($script:StepCount) : $Name"
        Write-Log -Level Step -Message ("=" * 70)
    }

    function Complete-Step {
        if ($null -eq $script:StepStart) {
            return
        }
        $duration = (Get-Date) - $script:StepStart
        Write-Log -Level Success -Message ("Fin de l'etape '{0}' en {1:hh\:mm\:ss}" -f $script:StepName, $duration)
        $script:StepStart = $null
    }

    function Write-ExecutionSummary {
        Complete-Step
        $total = (Get-Date) - $script:StartTime

        Write-Log -Level Step -Message ''
        Write-Log -Level Step -Message ("=" * 70)
        Write-Log -Level Step -Message ("RESUME - duree totale {0:hh\:mm\:ss} - mode {1}" -f $total, $(if ($Apply) { 'APPLICATION' } else { 'SIMULATION' }))
        Write-Log -Level Step -Message ("=" * 70)

        $groups = $script:MigrationLog | Group-Object Stage, Status | Sort-Object Name
        foreach ($group in $groups) {
            Write-Log -Level Info -Message ("{0,-40} {1,6}" -f $group.Name, $group.Count)
        }

        $errors = @($script:MigrationLog | Where-Object Status -eq 'Error')
        Write-Log -Level $(if ($errors.Count -gt 0) { 'Error' } else { 'Success' }) -Message "Erreurs : $($errors.Count)"
        foreach ($errorEntry in $errors | Select-Object -First 20) {
            Write-Log -Level Error -Message "  - [$($errorEntry.Stage)] $($errorEntry.Action) - $($errorEntry.Object) : $($errorEntry.Detail)"
        }
        if ($errors.Count -gt 20) {
            Write-Log -Level Error -Message "  ... $($errors.Count - 20) autre(s) erreur(s), voir MigrationLog.csv"
        }
        Write-Log -Level $(if ($script:Unresolved.Count -gt 0) { 'Warn' } else { 'Success' }) -Message "Principaux non resolus : $($script:Unresolved.Count) (voir UnresolvedPrincipals.csv)"
        if ($script:Invitations.Count -gt 0) {
            Write-Log -Level Info -Message "Invites crees : $($script:Invitations.Count) (voir Invitations.csv)"
        }

        $applied = @($script:MigrationLog | Where-Object { $_.Stage -eq 'Permissions' -and $_.Action -eq 'Autorisation ajoutee' }).Count
        $planned = @($script:MigrationLog | Where-Object { $_.Stage -eq 'Permissions' -and $_.Action -eq 'Ajouter autorisation' }).Count
        $notFound = @($script:MigrationLog | Where-Object { $_.Action -in @('Element cible introuvable', 'Objet cible introuvable') }).Count
        Write-Log -Level Info -Message "Autorisations appliquees : $applied - prevues (simulation) : $planned - objets cibles introuvables : $notFound"
        if (-not $Apply) {
            Write-Log -Level Warn -Message '>>> SIMULATION : RIEN N''A ETE MODIFIE SUR LE SITE CIBLE. Mettez $Apply = $true pour appliquer. <<<'
        }
    }

    function Add-MigrationLog {
        param(
            [Parameter(Mandatory)][string]$Stage,
            [Parameter(Mandatory)][string]$Action,
            [Parameter(Mandatory)][string]$Status,
            [string]$Object,
            [string]$Detail
        )

        $script:MigrationLog.Add([pscustomobject]@{
            Timestamp = (Get-Date).ToString('s')
            Stage     = $Stage
            Action    = $Action
            Status    = $Status
            Object    = $Object
            Detail    = $Detail
        })

        $level = switch ($Status) {
            'Error'   { 'Error' }
            'Skipped' { 'Warn' }
            'Planned' { 'Info' }
            default   { 'Success' }
        }
        $message = "[$Stage][$Status] $Action"
        if (-not [string]::IsNullOrWhiteSpace($Object)) { $message += " - $Object" }
        if (-not [string]::IsNullOrWhiteSpace($Detail)) { $message += " : $Detail" }
        Write-Log -Level $level -Message $message
    }

    function Add-UnresolvedPrincipal {
        param(
            [Parameter(Mandatory)]$Principal,
            [Parameter(Mandatory)][string]$Context,
            [Parameter(Mandatory)][string]$Reason
        )

        $script:Unresolved.Add([pscustomobject]@{
            SourceKey     = $Principal.Key
            Title         = $Principal.Title
            Email         = $Principal.Email
            LoginName     = $Principal.LoginName
            PrincipalType = $Principal.PrincipalType
            Category      = $Principal.Category
            Context       = $Context
            Reason        = $Reason
        })
        Add-MigrationLog -Stage 'Resolution' -Action 'Principal non resolu' -Status 'Skipped' -Object $Principal.Title -Detail $Reason
    }

    function Add-SkippedPrincipalOnce {
        param(
            [Parameter(Mandatory)]$Principal,
            [Parameter(Mandatory)][string]$Reason
        )

        if ($script:SkippedPrincipalKeys.ContainsKey([string]$Principal.Key)) {
            return
        }
        $script:SkippedPrincipalKeys[[string]$Principal.Key] = $true
        Add-MigrationLog -Stage 'Resolution' -Action 'Principal natif ignore' -Status 'Skipped' -Object $Principal.Title -Detail $Reason
    }

    function Test-IsGuid {
        param([string]$Value)
        $parsed = [guid]::Empty
        return [guid]::TryParse($Value, [ref]$parsed)
    }

    function Test-NativeRoleName {
        param([string]$Name)
        foreach ($group in $NativeRoleNameGroups) {
            if ($group -contains $Name) {
                return $true
            }
        }
        return $false
    }

    function Test-IsCustomRole {
        param([Parameter(Mandatory)]$Role)
        return ([string]$Role.RoleTypeKind -eq 'None' -and -not [bool]$Role.Hidden -and -not (Test-NativeRoleName -Name ([string]$Role.Name)))
    }

    function Get-PrincipalCategory {
        param(
            [string]$PrincipalType,
            [string]$LoginName,
            [string]$Title,
            [int]$Id
        )

        if ($PrincipalType -eq 'SharePointGroup') {
            if ($script:SourceAssociatedGroups.ContainsKey($Id)) {
                return 'AssociatedGroup'
            }
            if ($Title -like 'SharingLinks.*' -or $LoginName -like 'SharingLinks.*') {
                return 'SharingLink'
            }
            foreach ($pattern in $ExcludedGroupTitlePatterns) {
                if ($Title -like $pattern -or $LoginName -like $pattern) {
                    return 'System'
                }
            }
            return 'CustomGroup'
        }

        $login = $LoginName.ToLowerInvariant()
        if ($login -eq 'c:0(.s|true') { return 'Everyone' }
        if ($login -like 'c:0-.f|rolemanager|spo-grid-all-users*') { return 'EveryoneExceptExternal' }
        if ($login -like 'c:0t.c|tenant|*') { return 'EntraGroup' }
        if (-not [string]::IsNullOrWhiteSpace($script:SourceSiteGroupId) -and
            $login -like "c:0o.c|federateddirectoryclaimprovider|$($script:SourceSiteGroupId)*") {
            # Groupe Microsoft 365 du site source lui-meme (proprietaires "_o" ou membres) : natif.
            return 'SiteM365Group'
        }
        if ($login -like 'c:0o.c|federateddirectoryclaimprovider|*') { return 'M365Group' }
        if ($login -in @('sharepoint\system', 'c:0!.s|windows') -or
            $login -like 'nt authority\*' -or
            $login -like 'i:0i.t|ms.sp.ext|*' -or
            $login -like '*app@sharepoint') {
            return 'System'
        }
        if ($PrincipalType -eq 'User') { return 'User' }
        if ($PrincipalType -in @('SecurityGroup', 'DistributionList')) { return 'EntraGroup' }
        return 'System'
    }

    function Get-SourceUserEmail {
        param(
            [Parameter(Mandatory)]$Principal,
            [Parameter(Mandatory)]$Connection
        )

        $loginName = [string]$Principal.LoginName
        if ($script:SourceEmailCache.ContainsKey($loginName)) {
            return $script:SourceEmailCache[$loginName]
        }

        $email = ''
        try {
            # RoleAssignment.Member est un Principal : l'e-mail n'est disponible que sur l'objet User.
            if ($Principal -is [Microsoft.SharePoint.Client.User]) {
                Get-PnPProperty -ClientObject $Principal -Property Email -Connection $Connection | Out-Null
                $email = [string]$Principal.Email
            }
            else {
                $user = Get-PnPUser -Identity ([int]$Principal.Id) -Connection $Connection
                $email = [string]$user.Email
            }
        }
        catch {
            # Certains principaux n'exposent pas l'adresse e-mail.
        }

        if ([string]::IsNullOrWhiteSpace($email)) {
            $loginCandidate = ($loginName -split '\|')[-1]
            if ($loginCandidate -match '^urn(%3a|:)spo(%3a|:)guest#(?<mail>[^@\s]+@[^@\s]+)$') {
                # Externe SharePoint (code a usage unique) : urn:spo:guest#john@gmail.com
                $email = $Matches['mail']
            }
            elseif ($loginCandidate -match '^(?<local>.+)#ext#@') {
                # Invite du tenant source : john_gmail.com#ext#@... -> john@gmail.com
                $local = $Matches['local']
                $lastUnderscore = $local.LastIndexOf('_')
                if ($lastUnderscore -gt 0) {
                    $email = $local.Substring(0, $lastUnderscore) + '@' + $local.Substring($lastUnderscore + 1)
                }
            }
            elseif ($loginCandidate -match '^[^@\s]+@[^@\s]+$') {
                $email = $loginCandidate
            }
        }

        $email = $email.Trim().ToLowerInvariant()
        $script:SourceEmailCache[$loginName] = $email
        return $email
    }

    function Get-PrincipalRecord {
        param(
            [Parameter(Mandatory)]$Principal,
            [Parameter(Mandatory)]$Connection
        )

        $principalType = [string]$Principal.PrincipalType
        $loginName = [string]$Principal.LoginName
        $title = [string]$Principal.Title
        $id = [int]$Principal.Id
        $category = Get-PrincipalCategory -PrincipalType $principalType -LoginName $loginName -Title $title -Id $id

        $email = ''
        if ($category -eq 'User') {
            $email = Get-SourceUserEmail -Principal $Principal -Connection $Connection
        }

        $association = 'None'
        if ($category -eq 'AssociatedGroup') {
            $association = $script:SourceAssociatedGroups[$id]
        }

        # Externe = invite Entra (#ext#) ou externe SharePoint (urn:spo:guest).
        $loginLower = $loginName.ToLowerInvariant()
        $isExternal = ($category -eq 'User' -and ($loginLower -like '*#ext#*' -or $loginLower -like '*urn%3aspo%3aguest*' -or $loginLower -like '*urn:spo:guest*'))

        if ($category -eq 'User' -and -not [string]::IsNullOrWhiteSpace($email)) {
            $key = 'USER:' + $email
        }
        elseif ($principalType -eq 'SharePointGroup') {
            $key = 'SPGROUP:' + $title
        }
        elseif (-not [string]::IsNullOrWhiteSpace($loginName)) {
            $key = 'LOGIN:' + $loginName
        }
        else {
            $key = 'TITLE:' + $title
        }

        [pscustomobject]@{
            Key           = $key
            Title         = $title
            Email         = $email
            LoginName     = $loginName
            PrincipalType = $principalType
            Category      = $category
            Association   = $association
            IsExternal    = $isExternal
        }
    }

    function Get-RoleRecord {
        param(
            [Parameter(Mandatory)]$Role,
            [Parameter(Mandatory)]$Connection
        )

        Get-PnPProperty -ClientObject $Role -Property BasePermissions, Hidden -Connection $Connection | Out-Null
        $permissionNames = [System.Collections.Generic.List[string]]::new()

        foreach ($permissionKind in [Enum]::GetValues([Microsoft.SharePoint.Client.PermissionKind])) {
            if ($permissionKind -in @(
                    [Microsoft.SharePoint.Client.PermissionKind]::EmptyMask,
                    [Microsoft.SharePoint.Client.PermissionKind]::FullMask
                )) {
                continue
            }
            if ($Role.BasePermissions.Has($permissionKind)) {
                $permissionNames.Add([string]$permissionKind)
            }
        }

        [pscustomobject]@{
            Name            = [string]$Role.Name
            Description     = [string]$Role.Description
            RoleTypeKind    = [string]$Role.RoleTypeKind
            Hidden          = [bool]$Role.Hidden
            IsCustom        = (Test-IsCustomRole -Role $Role)
            PermissionKinds = @($permissionNames)
        }
    }

    function Get-Assignments {
        param(
            [Parameter(Mandatory)]$SecurableObject,
            [Parameter(Mandatory)]$Connection
        )

        Get-PnPProperty -ClientObject $SecurableObject -Property RoleAssignments -Connection $Connection | Out-Null
        $results = [System.Collections.Generic.List[object]]::new()

        foreach ($roleAssignment in $SecurableObject.RoleAssignments) {
            Get-PnPProperty -ClientObject $roleAssignment -Property Member, RoleDefinitionBindings -Connection $Connection | Out-Null
            $principal = Get-PrincipalRecord -Principal $roleAssignment.Member -Connection $Connection
            $roles = [System.Collections.Generic.List[object]]::new()

            foreach ($role in $roleAssignment.RoleDefinitionBindings) {
                # Les niveaux caches techniques (Acces limite, System.LimitedView/Edit...) sont
                # geres par SharePoint lui-meme. EXCEPTION : "Affichage restreint"
                # (RestrictedReader, cache) est un vrai droit accorde par un partage
                # "peut afficher, sans telechargement" : il doit etre migre.
                if ([bool]$role.Hidden -and [string]$role.RoleTypeKind -ne 'RestrictedReader') {
                    continue
                }
                $roles.Add([pscustomobject]@{
                    Name         = [string]$role.Name
                    RoleTypeKind = [string]$role.RoleTypeKind
                    IsCustom     = (Test-IsCustomRole -Role $role)
                })
            }

            if ($roles.Count -gt 0) {
                $results.Add([pscustomobject]@{
                    Principal = $principal
                    Roles     = @($roles)
                })
            }
        }
        return @($results)
    }

    function Get-AssociatedGroups {
        param(
            [Parameter(Mandatory)]$Connection,
            [Parameter(Mandatory)][string]$Stage
        )

        $groups = @{ Owners = $null; Members = $null; Visitors = $null }
        $associatedSwitches = @{
            Owners   = 'AssociatedOwnerGroup'
            Members  = 'AssociatedMemberGroup'
            Visitors = 'AssociatedVisitorGroup'
        }
        foreach ($association in $associatedSwitches.Keys) {
            try {
                $parameters = @{ Connection = $Connection }
                $parameters[$associatedSwitches[$association]] = $true
                $group = Get-PnPGroup @parameters
                if ($null -ne $group) {
                    $groups[$association] = $group
                }
            }
            catch {
                Add-MigrationLog -Stage $Stage -Action 'Lecture groupe associe' -Status 'Skipped' -Object $association -Detail $_.Exception.Message
            }
        }
        return $groups
    }

    function Get-ListItemSecurityRows {
        param(
            [Parameter(Mandatory)]$List,
            [Parameter(Mandatory)]$Connection
        )

        $listId = $List.Id.ToString()
        $nextUrl = "/_api/web/lists(guid'$listId')/items?`$select=Id,FileRef,FileSystemObjectType,HasUniqueRoleAssignments&`$top=5000"

        while (-not [string]::IsNullOrWhiteSpace($nextUrl)) {
            $page = Invoke-PnPSPRestMethod -Method Get -Url $nextUrl -Connection $Connection
            foreach ($row in @($page.value)) {
                if ([bool]$row.HasUniqueRoleAssignments) {
                    $row
                }
            }

            $nextLink = $null
            if ($null -ne $page.PSObject.Properties['@odata.nextLink']) {
                $nextLink = $page.PSObject.Properties['@odata.nextLink'].Value
            }
            elseif ($null -ne $page.PSObject.Properties['odata.nextLink']) {
                $nextLink = $page.PSObject.Properties['odata.nextLink'].Value
            }
            if ([string]::IsNullOrWhiteSpace([string]$nextLink)) {
                $nextUrl = $null
            }
            elseif ([Uri]::IsWellFormedUriString([string]$nextLink, [UriKind]::Absolute)) {
                $nextUrl = ([Uri]$nextLink).PathAndQuery
            }
            else {
                $nextUrl = [string]$nextLink
            }
        }
    }

    function Export-SharePointSecurity {
        param(
            [Parameter(Mandatory)]$Connection,
            [Parameter(Mandatory)][string]$ExportFile
        )

        Write-Log -Message "Lecture du site source : $SourceSiteUrl"
        $script:SourceSiteGroupId = Get-SiteGroupId -Connection $Connection
        if ($SharingLinkRecipients -ne 'None') {
            try {
                $siteUri = [Uri]$SourceSiteUrl
                $graphSite = Invoke-PnPGraphMethod -Method Get -Url "sites/$($siteUri.Host):$($siteUri.AbsolutePath.TrimEnd('/'))?`$select=id" -Connection $Connection
                $script:SourceGraphSiteId = [string]$graphSite.id
                Write-Log -Message "Lecture des partages via Microsoft Graph active (site $($script:SourceGraphSiteId))."
            }
            catch {
                $script:SourceGraphSiteId = $null
                Write-Log -Level Warn -Message "Lecture des partages via Graph impossible : $($_.Exception.Message)"
                Write-Log -Level Warn -Message '  -> Ajoutez la permission Graph APPLICATION "Sites.Read.All" (ou Files.Read.All) a l''application source, avec consentement administrateur.'
            }
        }
        Write-Log -Message "Groupe Microsoft 365 du site source : $(if ($script:SourceSiteGroupId) { $script:SourceSiteGroupId } else { '(site non connecte a un groupe)' })"
        $web = Get-PnPWeb -Includes ServerRelativeUrl, RoleAssignments, Title -Connection $Connection
        $webRelativeUrl = [string]$web.ServerRelativeUrl
        Add-MigrationLog -Stage 'Export' -Action 'Site source lu' -Status 'Success' -Object $web.Title -Detail $webRelativeUrl

        Write-Log -Message 'Lecture des niveaux d''autorisation...'
        $roleDefinitions = [System.Collections.Generic.List[object]]::new()
        foreach ($role in Get-PnPRoleDefinition -Connection $Connection) {
            $roleRecord = Get-RoleRecord -Role $role -Connection $Connection
            if ($roleRecord.IsCustom) {
                $roleDefinitions.Add($roleRecord)
                Write-Log -Level Info -Message "  Niveau PERSONNALISE : '$($roleRecord.Name)' ($(@($roleRecord.PermissionKinds).Count) permissions)"
            }
            else {
                Write-Log -Level Detail -Message "  Niveau natif ignore : '$($roleRecord.Name)' (RoleTypeKind=$($roleRecord.RoleTypeKind), Hidden=$($roleRecord.Hidden))"
            }
        }
        Write-Log -Level Success -Message "$($roleDefinitions.Count) niveau(x) personnalise(s) a migrer."

        Write-Log -Message 'Lecture des groupes natifs du site (Proprietaires/Membres/Visiteurs)...'
        $associatedGroups = Get-AssociatedGroups -Connection $Connection -Stage 'Export'
        foreach ($association in $associatedGroups.Keys) {
            if ($null -ne $associatedGroups[$association]) {
                $script:SourceAssociatedGroups[[int]$associatedGroups[$association].Id] = $association
                Write-Log -Level Detail -Message "  $association = '$($associatedGroups[$association].Title)' (ID $($associatedGroups[$association].Id))"
            }
        }

        Write-Log -Message 'Lecture des groupes SharePoint...'
        $allGroups = @(Get-PnPGroup -Connection $Connection)
        Write-Log -Level Detail -Message "  $($allGroups.Count) groupe(s) trouve(s) sur le site source."
        $groups = [System.Collections.Generic.List[object]]::new()
        $sharingLinkGroups = [System.Collections.Generic.List[object]]::new()
        foreach ($group in $allGroups) {
            $category = Get-PrincipalCategory -PrincipalType 'SharePointGroup' -LoginName ([string]$group.LoginName) -Title ([string]$group.Title) -Id ([int]$group.Id)
            if ($category -eq 'SharingLink') {
                if ($SharingLinkRecipients -eq 'None') {
                    Add-MigrationLog -Stage 'Export' -Action 'Lien de partage ignore (SharingLinkRecipients = None)' -Status 'Skipped' -Object $group.Title
                    continue
                }
                # SharingLinks.<ID objet>.<Type>.<ID partage> : Flexible = personnes specifiques,
                # AnonymousView/Edit = tout le monde, OrganizationView/Edit = organisation.
                $titleParts = ([string]$group.Title).Split('.')
                $linkType = if ($titleParts.Count -ge 3) { $titleParts[2] } else { 'Inconnu' }
                $linkMembers = [System.Collections.Generic.List[object]]::new()
                try {
                    foreach ($member in Get-PnPGroupMember -Group $group -Connection $Connection) {
                        $linkMembers.Add((Get-PrincipalRecord -Principal $member -Connection $Connection))
                    }
                }
                catch {
                    Add-MigrationLog -Stage 'Export' -Action 'Lecture destinataires du lien' -Status 'Error' -Object $group.Title -Detail $_.Exception.Message
                }
                $sharingLinkGroups.Add([pscustomobject]@{
                    Title    = [string]$group.Title
                    LinkType = $linkType
                    Members  = @($linkMembers)
                })
                $externalCount = @($linkMembers | Where-Object { [bool]$_.IsExternal }).Count
                Write-Log -Level Info -Message "  Lien de partage ($linkType) : $($linkMembers.Count) destinataire(s) dont $externalCount externe(s)"
                foreach ($member in $linkMembers) {
                    Write-Log -Level Detail -Message "    destinataire : $($member.Title) <$($member.Email)> $(if ($member.IsExternal) { '[EXTERNE]' } else { '[interne]' })"
                }
                continue
            }
            if ($category -eq 'AssociatedGroup') {
                Add-MigrationLog -Stage 'Export' -Action 'Groupe natif du site non recree' -Status 'Skipped' -Object $group.Title
                continue
            }
            if ($category -eq 'System') {
                Add-MigrationLog -Stage 'Export' -Action 'Groupe systeme/partage non recree' -Status 'Skipped' -Object $group.Title
                continue
            }

            Get-PnPProperty -ClientObject $group -Property Description, Owner, AllowMembersEditMembership, OnlyAllowMembersViewMembership, AllowRequestToJoinLeave, AutoAcceptRequestToJoinLeave, RequestToJoinLeaveEmailSetting -Connection $Connection | Out-Null

            $owner = $null
            try {
                $owner = Get-PrincipalRecord -Principal $group.Owner -Connection $Connection
            }
            catch {
                Add-MigrationLog -Stage 'Export' -Action 'Lecture proprietaire du groupe' -Status 'Skipped' -Object $group.Title -Detail $_.Exception.Message
            }

            $members = [System.Collections.Generic.List[object]]::new()
            try {
                foreach ($member in Get-PnPGroupMember -Group $group -Connection $Connection) {
                    $members.Add((Get-PrincipalRecord -Principal $member -Connection $Connection))
                }
            }
            catch {
                Add-MigrationLog -Stage 'Export' -Action 'Lecture membres du groupe' -Status 'Error' -Object $group.Title -Detail $_.Exception.Message
            }

            $groups.Add([pscustomobject]@{
                Title                          = [string]$group.Title
                Description                    = [string]$group.Description
                Owner                          = $owner
                AllowMembersEditMembership     = [bool]$group.AllowMembersEditMembership
                OnlyAllowMembersViewMembership = [bool]$group.OnlyAllowMembersViewMembership
                AllowRequestToJoinLeave        = [bool]$group.AllowRequestToJoinLeave
                AutoAcceptRequestToJoinLeave   = [bool]$group.AutoAcceptRequestToJoinLeave
                RequestToJoinLeaveEmailSetting = [string]$group.RequestToJoinLeaveEmailSetting
                Members                        = @($members)
            })
            $ownerTitle = if ($null -ne $owner) { $owner.Title } else { '?' }
            Write-Log -Level Info -Message "  Groupe PERSONNALISE : '$($group.Title)' - $($members.Count) membre(s), proprietaire '$ownerTitle'"
            foreach ($member in $members) {
                Write-Log -Level Detail -Message "    membre : $($member.Title) [$($member.Category)] $($member.Email)"
            }
        }
        Write-Log -Level Success -Message "$($groups.Count) groupe(s) personnalise(s) a migrer."

        Write-Log -Message 'Lecture des autorisations du site racine...'
        $securableObjects = [System.Collections.Generic.List[object]]::new()
        $securableObjects.Add([pscustomobject]@{
            ObjectType      = 'Web'
            RelativePath    = ''
            ListRelativeUrl = $null
            IsLibrary       = $false
            ItemId          = $null
            FileRef         = $null
            IsFolder        = $false
            Assignments     = @(Get-Assignments -SecurableObject $web -Connection $Connection)
            SharingGrants   = @()
        })

        Write-Log -Level Detail -Message "  $(@($securableObjects[0].Assignments).Count) attribution(s) au niveau du site."

        Write-Log -Message 'Lecture des listes et bibliotheques...'
        $lists = @(Get-PnPList -Includes RootFolder, Hidden, IsSystemList, HasUniqueRoleAssignments, BaseType, ItemCount -Connection $Connection)
        Write-Log -Level Info -Message "  $($lists.Count) liste(s) trouvee(s) (y compris cachees/systeme)."
        $listIndex = 0
        foreach ($list in $lists) {
            $listIndex++
            Show-Progress -Activity 'Export des autorisations SharePoint' -Status ([string]$list.Title) -Index $listIndex -Total $lists.Count

            if (-not $IncludeHiddenLists -and ([bool]$list.Hidden -or [bool]$list.IsSystemList)) {
                Write-Log -Level Detail -Message "  [$listIndex/$($lists.Count)] Liste cachee/systeme ignoree : $($list.Title)"
                continue
            }

            $listRelativeUrl = ([string]$list.RootFolder.ServerRelativeUrl).Substring($webRelativeUrl.Length).TrimStart('/')
            $isLibrary = ([string]$list.BaseType -eq 'DocumentLibrary')
            $listStart = Get-Date
            $objectsBefore = $securableObjects.Count
            Write-Log -Level Info -Message "  [$listIndex/$($lists.Count)] Analyse de '$($list.Title)' ($listRelativeUrl) - $($list.ItemCount) element(s), droits uniques : $([bool]$list.HasUniqueRoleAssignments)"

            try {
                if ([bool]$list.HasUniqueRoleAssignments) {
                    $securableObjects.Add([pscustomobject]@{
                        ObjectType      = 'List'
                        RelativePath    = $listRelativeUrl
                        ListRelativeUrl = $listRelativeUrl
                        IsLibrary       = $isLibrary
                        ItemId          = $null
                        FileRef         = $null
                        IsFolder        = $false
                        Assignments     = @(Get-Assignments -SecurableObject $list -Connection $Connection)
                        SharingGrants   = @()
                    })
                }

                if ($SkipItemPermissions -or $list.ItemCount -eq 0) {
                    continue
                }

                foreach ($row in Get-ListItemSecurityRows -List $list -Connection $Connection) {
                    try {
                        $item = Get-PnPListItem -List $list -Id ([int]$row.Id) -Connection $Connection
                        $fileRef = [string]$row.FileRef
                        $relativePath = if (-not [string]::IsNullOrWhiteSpace($fileRef)) {
                            $fileRef.Substring($webRelativeUrl.Length).TrimStart('/')
                        }
                        else {
                            "$listRelativeUrl|ID:$($row.Id)"
                        }

                        # Partages (liens "personnes specifiques" et invitations en attente) lus via Graph.
                        $sharingGrants = @()
                        Write-Log -Level Detail -Message "      Objet : $relativePath"
                        if ($isLibrary -and $SharingLinkRecipients -ne 'None') {
                            $sharingGrants = @(Get-SourceSharingGrants -Connection $Connection -ListId $list.Id.ToString() -ItemId ([int]$row.Id) -Context $relativePath)
                        }

                        $securableObjects.Add([pscustomobject]@{
                            ObjectType      = 'Item'
                            RelativePath    = $relativePath
                            ListRelativeUrl = $listRelativeUrl
                            IsLibrary       = $isLibrary
                            ItemId          = [int]$row.Id
                            FileRef         = $fileRef
                            IsFolder        = ([int]$row.FileSystemObjectType -eq 1)
                            Assignments     = @(Get-Assignments -SecurableObject $item -Connection $Connection)
                            SharingGrants   = $sharingGrants
                        })
                        Write-Log -Level Detail -Message "      Droits uniques : $relativePath ($(@($securableObjects[$securableObjects.Count - 1].Assignments).Count) attribution(s), $($sharingGrants.Count) partage(s))"
                    }
                    catch {
                        Add-MigrationLog -Stage 'Export' -Action 'Lecture droits element' -Status 'Error' -Object "$listRelativeUrl / ID $($row.Id)" -Detail $_.Exception.Message
                    }
                }
            }
            catch {
                Add-MigrationLog -Stage 'Export' -Action 'Lecture droits liste' -Status 'Error' -Object $listRelativeUrl -Detail $_.Exception.Message
            }
            finally {
                Write-Log -Level Detail -Message ("      -> {0} objet(s) a droits uniques dans cette liste, en {1:N1} s" -f ($securableObjects.Count - $objectsBefore), ((Get-Date) - $listStart).TotalSeconds)
            }
        }
        Write-Progress -Activity 'Export des autorisations SharePoint' -Completed

        $export = [pscustomobject]@{
            SchemaVersion        = 2
            ExportedAtUtc        = (Get-Date).ToUniversalTime().ToString('o')
            SourceSiteUrl        = $SourceSiteUrl
            SourceWebRelativeUrl = $webRelativeUrl
            CustomRoles          = @($roleDefinitions)
            SharePointGroups     = @($groups)
            SecurableObjects     = @($securableObjects)
            SharingLinkGroups    = @($sharingLinkGroups)
        }

        $export | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $ExportFile -Encoding utf8
        Add-MigrationLog -Stage 'Export' -Action 'Export JSON termine' -Status 'Success' -Object $ExportFile -Detail "$($roleDefinitions.Count) niveaux, $($groups.Count) groupes, $($securableObjects.Count) objets securises"
        return $export
    }

    function Import-PrincipalMappings {
        if ([string]::IsNullOrWhiteSpace($PrincipalMappingCsv)) {
            return
        }
        if (-not (Test-Path -LiteralPath $PrincipalMappingCsv)) {
            throw "Fichier de correspondance introuvable : $PrincipalMappingCsv"
        }

        foreach ($mapping in Import-Csv -LiteralPath $PrincipalMappingCsv) {
            if ([string]::IsNullOrWhiteSpace([string]$mapping.SourceKey) -or [string]::IsNullOrWhiteSpace([string]$mapping.TargetLogin)) {
                throw 'PrincipalMappingCsv doit contenir les colonnes SourceKey et TargetLogin, sans valeur vide.'
            }
            $script:PrincipalMappings[[string]$mapping.SourceKey] = [string]$mapping.TargetLogin
        }
    }

    function Get-CandidateEmails {
        param([Parameter(Mandatory)][string]$Email)

        $candidates = [System.Collections.Generic.List[string]]::new()
        $parts = $Email.Split('@')
        if ($parts.Count -eq 2 -and $DomainMapping.ContainsKey($parts[1])) {
            $candidates.Add(($parts[0] + '@' + [string]$DomainMapping[$parts[1]]).ToLowerInvariant())
        }
        if (-not $candidates.Contains($Email)) {
            $candidates.Add($Email)
        }
        return @($candidates)
    }

    function Get-GraphUserByEmail {
        param(
            [Parameter(Mandatory)][string]$Email,
            [Parameter(Mandatory)]$Connection
        )

        $escapedEmail = $Email.Replace("'", "''")
        $filter = [Uri]::EscapeDataString("mail eq '$escapedEmail' or userPrincipalName eq '$escapedEmail'")
        $response = Invoke-PnPGraphMethod -Method Get -Url "users?`$select=id,displayName,mail,userPrincipalName,userType&`$filter=$filter" -Connection $Connection
        $user = @($response.value) | Select-Object -First 1
        if ($null -ne $user) {
            return $user
        }

        # Alias SMTP (comptes migres) ou adresse d'origine d'un invite (otherMails).
        # Requete avancee Graph : indisponible avec les anciennes versions du module.
        if (-not $script:GraphSupportsEventual) {
            return $null
        }
        $advancedFilter = [Uri]::EscapeDataString("proxyAddresses/any(x:x eq 'smtp:$escapedEmail') or otherMails/any(x:x eq '$escapedEmail')")
        $response = Invoke-PnPGraphMethod -Method Get `
            -Url "users?`$select=id,displayName,mail,userPrincipalName,userType&`$count=true&`$filter=$advancedFilter" `
            -ConsistencyLevelEventual -Connection $Connection
        return @($response.value) | Select-Object -First 1
    }

    function Confirm-TargetLogin {
        param(
            [Parameter(Mandatory)][string]$Login,
            [Parameter(Mandatory)]$Principal,
            [Parameter(Mandatory)]$Connection,
            [Parameter(Mandatory)][string]$Context,
            # Echec silencieux (une autre methode de resolution sera tentee ensuite).
            [switch]$Quiet
        )

        if (-not $Apply) {
            return $true
        }

        $maxAttempts = if ($Quiet) { 1 } else { 6 }
        $lastError = $null
        for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
            try {
                New-PnPUser -LoginName $Login -Connection $Connection | Out-Null
                return $true
            }
            catch {
                $lastError = $_.Exception.Message
                if ($attempt -lt $maxAttempts) {
                    Start-Sleep -Seconds 5
                }
            }
        }
        if ($Quiet) {
            Write-Log -Level Detail -Message "    '$Login' non resolu par SharePoint ($lastError), autre methode tentee."
            return $false
        }
        Add-UnresolvedPrincipal -Principal $Principal -Context $Context -Reason "Present dans Entra mais non resolu par SharePoint (partage externe du site ?) : $lastError"
        return $false
    }

    function Resolve-TargetUser {
        param(
            [Parameter(Mandatory)]$Principal,
            [Parameter(Mandatory)]$Connection,
            [Parameter(Mandatory)][string]$Context
        )

        $key = [string]$Principal.Key
        if ($script:TargetPrincipalCache.ContainsKey($key)) {
            return $script:TargetPrincipalCache[$key]
        }
        if ($script:PrincipalMappings.ContainsKey($key)) {
            $mapped = $script:PrincipalMappings[$key]
            Write-Log -Level Detail -Message "  Utilisateur '$($Principal.Title)' -> '$mapped' (fichier de correspondance CSV)"
            if (-not (Confirm-TargetLogin -Login $mapped -Principal $Principal -Connection $Connection -Context $Context)) {
                return $null
            }
            $script:TargetPrincipalCache[$key] = $mapped
            return $mapped
        }

        # Meme tenant : le compte est identique des deux cotes, on reutilise son login.
        if ($script:SameTenant -and -not [string]::IsNullOrWhiteSpace([string]$Principal.LoginName)) {
            $sameLogin = [string]$Principal.LoginName
            $isExternal = [bool]$Principal.IsExternal
            # Un externe SharePoint (urn:spo:guest) peut ne pas etre resolvable sur un autre site :
            # dans ce cas on passe par Graph / invitation B2B avec son adresse e-mail.
            if (Confirm-TargetLogin -Login $sameLogin -Principal $Principal -Connection $Connection -Context $Context -Quiet:$isExternal) {
                Write-Log -Level Detail -Message "  Utilisateur (meme tenant) : '$($Principal.Title)' -> $sameLogin"
                $script:TargetPrincipalCache[$key] = $sameLogin
                return $sameLogin
            }
            if (-not $isExternal) {
                $script:TargetPrincipalCache[$key] = $null
                return $null
            }
        }

        $email = [string]$Principal.Email
        if ([string]::IsNullOrWhiteSpace($email)) {
            Add-UnresolvedPrincipal -Principal $Principal -Context $Context -Reason 'Utilisateur sans adresse e-mail. Ajouter une correspondance CSV.'
            return $null
        }

        $targetUser = $null
        try {
            foreach ($candidate in Get-CandidateEmails -Email $email) {
                Write-Log -Level Detail -Message "  Recherche Graph de '$candidate' dans le tenant cible..."
                $targetUser = Get-GraphUserByEmail -Email $candidate -Connection $Connection
                if ($null -ne $targetUser) {
                    Write-Log -Level Detail -Message "    trouve : $($targetUser.userPrincipalName) ($($targetUser.userType))"
                    break
                }
                Write-Log -Level Detail -Message '    introuvable.'
            }
        }
        catch {
            Add-UnresolvedPrincipal -Principal $Principal -Context $Context -Reason "Recherche Graph impossible : $($_.Exception.Message)"
            return $null
        }

        if ($null -eq $targetUser) {
            $canInvite = $InviteMissingUsers -or ([bool]$Principal.IsExternal -and $InviteExternalRecipients)
            if (-not $canInvite) {
                Add-UnresolvedPrincipal -Principal $Principal -Context $Context -Reason 'Absent du tenant cible. Renseigner DomainMapping, une correspondance CSV, ou activer InviteMissingUsers / InviteExternalRecipients.'
                $script:TargetPrincipalCache[$key] = $null
                return $null
            }

            if (-not $Apply) {
                Add-MigrationLog -Stage 'Invitation' -Action 'Creer invite sans e-mail' -Status 'Planned' -Object $email -Detail 'sendInvitationMessage=false'
                $script:TargetPrincipalCache[$key] = $email
                return $email
            }

            try {
                $invitation = Invoke-PnPGraphMethod -Method Post -Url 'invitations' -Content @{
                    invitedUserEmailAddress = $email
                    inviteRedirectUrl       = $TargetSiteUrl
                    sendInvitationMessage   = $false
                } -Connection $Connection

                # La reponse ne contient que l'id : on relit l'UPN (replication Entra).
                $invitedId = [string]$invitation.invitedUser.id
                $targetUser = $null
                for ($attempt = 1; $attempt -le 6 -and $null -eq $targetUser; $attempt++) {
                    try {
                        $targetUser = Invoke-PnPGraphMethod -Method Get -Url "users/$($invitedId)?`$select=id,userPrincipalName,mail" -Connection $Connection
                    }
                    catch {
                        Start-Sleep -Seconds 5
                    }
                }
                if ($null -eq $targetUser) {
                    throw "Invite $invitedId cree mais introuvable dans Graph."
                }

                $script:Invitations.Add([pscustomobject]@{
                    Email             = $email
                    UserPrincipalName = [string]$targetUser.userPrincipalName
                    Status            = [string]$invitation.status
                    RedeemUrl         = [string]$invitation.inviteRedeemUrl
                })
                Add-MigrationLog -Stage 'Invitation' -Action 'Invite cree sans e-mail' -Status 'Success' -Object $email -Detail 'Le lien de redemption est conserve dans Invitations.csv.'
            }
            catch {
                Add-UnresolvedPrincipal -Principal $Principal -Context $Context -Reason "Creation de l'invite impossible : $($_.Exception.Message)"
                $script:TargetPrincipalCache[$key] = $null
                return $null
            }
        }

        $login = 'i:0#.f|membership|' + ([string]$targetUser.userPrincipalName).ToLowerInvariant()
        if (-not (Confirm-TargetLogin -Login $login -Principal $Principal -Connection $Connection -Context $Context)) {
            $script:TargetPrincipalCache[$key] = $null
            return $null
        }

        Write-Log -Level Info -Message "  Utilisateur resolu : '$($Principal.Title)' ($email) -> $login"
        $script:TargetPrincipalCache[$key] = $login
        return $login
    }

    function Resolve-TargetEntraGroup {
        param(
            [Parameter(Mandatory)]$Principal,
            [Parameter(Mandatory)]$Connection,
            [Parameter(Mandatory)][string]$Context
        )

        $key = [string]$Principal.Key
        if ($script:TargetPrincipalCache.ContainsKey($key)) {
            return $script:TargetPrincipalCache[$key]
        }

        $login = $null
        if ($script:PrincipalMappings.ContainsKey($key)) {
            $login = $script:PrincipalMappings[$key]
        }
        elseif ($script:SameTenant) {
            # Meme tenant : le groupe Entra/M365 est identique, on reutilise sa claim.
            $login = [string]$Principal.LoginName
        }
        else {
            try {
                Write-Log -Level Detail -Message "  Recherche Graph du groupe Entra '$($Principal.Title)' dans le tenant cible..."
                $escapedName = ([string]$Principal.Title).Replace("'", "''")
                $filter = [Uri]::EscapeDataString("displayName eq '$escapedName'")
                $response = Invoke-PnPGraphMethod -Method Get -Url "groups?`$select=id,displayName&`$filter=$filter" -Connection $Connection
                $found = @($response.value)
            }
            catch {
                Add-UnresolvedPrincipal -Principal $Principal -Context $Context -Reason "Recherche du groupe Entra impossible : $($_.Exception.Message)"
                $script:TargetPrincipalCache[$key] = $null
                return $null
            }

            if ($found.Count -ne 1) {
                Add-UnresolvedPrincipal -Principal $Principal -Context $Context -Reason "Groupe Entra '$($Principal.Title)' : $($found.Count) correspondance(s) dans le tenant cible. Ajouter une correspondance CSV."
                $script:TargetPrincipalCache[$key] = $null
                return $null
            }

            $groupId = [string]$found[0].id
            if ([string]$Principal.Category -eq 'M365Group') {
                # Suffixe _o = proprietaires du groupe Microsoft 365.
                $suffix = if (([string]$Principal.LoginName).EndsWith('_o')) { '_o' } else { '' }
                $login = "c:0o.c|federateddirectoryclaimprovider|$groupId$suffix"
            }
            else {
                $login = "c:0t.c|tenant|$groupId"
            }
        }

        if (-not (Confirm-TargetLogin -Login $login -Principal $Principal -Connection $Connection -Context $Context)) {
            $script:TargetPrincipalCache[$key] = $null
            return $null
        }
        Write-Log -Level Info -Message "  Groupe Entra resolu : '$($Principal.Title)' -> $login"
        $script:TargetPrincipalCache[$key] = $login
        return $login
    }

    function Resolve-AssignmentPrincipal {
        param(
            [Parameter(Mandatory)]$Principal,
            [Parameter(Mandatory)]$Connection,
            [Parameter(Mandatory)][string]$Context,
            [Parameter(Mandatory)]$TargetAssociated
        )

        switch ([string]$Principal.Category) {
            'SharingLink' {
                # Les liens sont developpes en droits directs avant cet appel (Expand-SharingLinkAssignments).
                Add-SkippedPrincipalOnce -Principal $Principal -Reason 'Lien de partage : non recree (destinataires traites a part).'
                return $null
            }
            'System' {
                Add-SkippedPrincipalOnce -Principal $Principal -Reason 'Compte ou groupe systeme.'
                return $null
            }
            'Everyone' {
                if (-not $MigrateEveryoneClaims) {
                    Add-SkippedPrincipalOnce -Principal $Principal -Reason 'Tout le monde : MigrateEveryoneClaims = $false.'
                    return $null
                }
                return [pscustomobject]@{ Kind = 'User'; Value = 'c:0(.s|true' }
            }
            'EveryoneExceptExternal' {
                if (-not $MigrateEveryoneClaims -or [string]::IsNullOrWhiteSpace($script:TargetTenantGuid)) {
                    Add-SkippedPrincipalOnce -Principal $Principal -Reason 'Tout le monde sauf externes : MigrateEveryoneClaims = $false ou GUID du tenant cible inconnu.'
                    return $null
                }
                return [pscustomobject]@{ Kind = 'User'; Value = "c:0-.f|rolemanager|spo-grid-all-users/$($script:TargetTenantGuid)" }
            }
            'SiteM365Group' {
                # Groupe M365 du site source -> groupe M365 du site cible (meme role : proprietaires ou membres).
                if (-not $IncludeAssociatedGroupAssignments -or [string]::IsNullOrWhiteSpace($script:TargetSiteGroupId)) {
                    Add-SkippedPrincipalOnce -Principal $Principal -Reason 'Groupe Microsoft 365 du site source (site cible non connecte a un groupe).'
                    return $null
                }
                $suffix = if (([string]$Principal.LoginName).EndsWith('_o')) { '_o' } else { '' }
                $claim = "c:0o.c|federateddirectoryclaimprovider|$($script:TargetSiteGroupId)$suffix"
                Write-Log -Level Detail -Message "    groupe M365 du site : '$($Principal.Title)' -> groupe M365 du site cible ($claim)"
                return [pscustomobject]@{ Kind = 'User'; Value = $claim }
            }
            'AssociatedGroup' {
                $targetGroup = $TargetAssociated[[string]$Principal.Association]
                if (-not $IncludeAssociatedGroupAssignments -or $null -eq $targetGroup) {
                    Add-SkippedPrincipalOnce -Principal $Principal -Reason 'Groupe natif du site.'
                    return $null
                }
                return [pscustomobject]@{ Kind = 'Group'; Value = [string]$targetGroup.Title }
            }
            'CustomGroup' {
                if ($script:PrincipalMappings.ContainsKey([string]$Principal.Key)) {
                    return [pscustomobject]@{ Kind = 'Group'; Value = $script:PrincipalMappings[[string]$Principal.Key] }
                }
                if ($script:TargetGroupMap.ContainsKey([string]$Principal.Key)) {
                    return [pscustomobject]@{ Kind = 'Group'; Value = $script:TargetGroupMap[[string]$Principal.Key] }
                }
                Add-UnresolvedPrincipal -Principal $Principal -Context $Context -Reason 'Groupe SharePoint cible introuvable.'
                return $null
            }
            { $_ -in @('EntraGroup', 'M365Group') } {
                $login = Resolve-TargetEntraGroup -Principal $Principal -Connection $Connection -Context $Context
                if ([string]::IsNullOrWhiteSpace($login)) {
                    return $null
                }
                return [pscustomobject]@{ Kind = 'User'; Value = $login }
            }
            default {
                $login = Resolve-TargetUser -Principal $Principal -Connection $Connection -Context $Context
                if ([string]::IsNullOrWhiteSpace($login)) {
                    return $null
                }
                return [pscustomobject]@{ Kind = 'User'; Value = $login }
            }
        }
    }

    function Initialize-TargetRoles {
        param(
            [Parameter(Mandatory)]$Export,
            [Parameter(Mandatory)]$Connection
        )

        $targetRoles = @(Get-PnPRoleDefinition -Connection $Connection)
        foreach ($role in $targetRoles) {
            $script:TargetRoleCache["NAME:$($role.Name)"] = $role.Name
            if ([string]$role.RoleTypeKind -ne 'None') {
                $script:TargetRoleCache["TYPE:$($role.RoleTypeKind)"] = $role.Name
            }
        }

        foreach ($customRole in @($Export.CustomRoles)) {
            $existing = $targetRoles | Where-Object Name -eq $customRole.Name | Select-Object -First 1
            $wanted = (@($customRole.PermissionKinds) | Sort-Object) -join ','

            if ($null -eq $existing) {
                if ($Apply) {
                    $parameters = @{
                        RoleName    = $customRole.Name
                        Description = $customRole.Description
                        Connection  = $Connection
                    }
                    if (@($customRole.PermissionKinds).Count -gt 0) {
                        $parameters.Include = @($customRole.PermissionKinds)
                    }
                    Add-PnPRoleDefinition @parameters | Out-Null
                    Add-MigrationLog -Stage 'Roles' -Action 'Niveau personnalise cree' -Status 'Success' -Object $customRole.Name
                }
                else {
                    Add-MigrationLog -Stage 'Roles' -Action 'Creer niveau personnalise' -Status 'Planned' -Object $customRole.Name
                }
            }
            elseif ([string]$existing.RoleTypeKind -ne 'None' -or (Test-NativeRoleName -Name ([string]$existing.Name))) {
                # Ne jamais modifier un niveau natif du site cible.
                Add-MigrationLog -Stage 'Roles' -Action 'Nom reserve a un niveau natif cible, non modifie' -Status 'Skipped' -Object $customRole.Name
            }
            else {
                $current = (@((Get-RoleRecord -Role $existing -Connection $Connection).PermissionKinds) | Sort-Object) -join ','
                if ($current -eq $wanted -and [string]$existing.Description -eq [string]$customRole.Description) {
                    Add-MigrationLog -Stage 'Roles' -Action 'Niveau personnalise deja conforme' -Status 'Success' -Object $customRole.Name
                }
                elseif ($Apply) {
                    Set-PnPRoleDefinition -Identity $existing -Description $customRole.Description -ClearAll -Connection $Connection | Out-Null
                    if (@($customRole.PermissionKinds).Count -gt 0) {
                        Set-PnPRoleDefinition -Identity $existing -Select @($customRole.PermissionKinds) -Connection $Connection | Out-Null
                    }
                    Add-MigrationLog -Stage 'Roles' -Action 'Niveau personnalise synchronise' -Status 'Success' -Object $customRole.Name
                }
                else {
                    Add-MigrationLog -Stage 'Roles' -Action 'Synchroniser niveau personnalise' -Status 'Planned' -Object $customRole.Name -Detail "Actuel: $current | Voulu: $wanted"
                }
            }
            $script:TargetRoleCache["NAME:$($customRole.Name)"] = $customRole.Name
        }
    }

    function Resolve-TargetRoleName {
        param([Parameter(Mandatory)]$Role)

        $name = [string]$Role.Name
        if ([bool]$Role.IsCustom) {
            if ($script:TargetRoleCache.ContainsKey("NAME:$name")) {
                return $script:TargetRoleCache["NAME:$name"]
            }
            return $null
        }
        if ([string]$Role.RoleTypeKind -ne 'None' -and $script:TargetRoleCache.ContainsKey("TYPE:$($Role.RoleTypeKind)")) {
            return $script:TargetRoleCache["TYPE:$($Role.RoleTypeKind)"]
        }
        # Niveau natif sans RoleTypeKind : equivalence de nom selon la langue.
        foreach ($group in $NativeRoleNameGroups) {
            if ($group -contains $name) {
                foreach ($candidate in $group) {
                    if ($script:TargetRoleCache.ContainsKey("NAME:$candidate")) {
                        return $script:TargetRoleCache["NAME:$candidate"]
                    }
                }
            }
        }
        if ($script:TargetRoleCache.ContainsKey("NAME:$name")) {
            return $script:TargetRoleCache["NAME:$name"]
        }
        return $null
    }

    function Resolve-GroupOwner {
        param(
            [Parameter(Mandatory)]$Owner,
            [Parameter(Mandatory)]$Connection,
            [Parameter(Mandatory)]$TargetAssociated,
            [Parameter(Mandatory)][string]$Context
        )

        switch ([string]$Owner.Category) {
            'AssociatedGroup' {
                $group = $TargetAssociated[[string]$Owner.Association]
                if ($null -ne $group) { return [string]$group.Title }
                return $null
            }
            'CustomGroup' {
                if ($script:TargetGroupMap.ContainsKey([string]$Owner.Key)) {
                    return $script:TargetGroupMap[[string]$Owner.Key]
                }
                return $null
            }
            'User' {
                return Resolve-TargetUser -Principal $Owner -Connection $Connection -Context $Context
            }
            default {
                return $null
            }
        }
    }

    function Initialize-TargetGroups {
        param(
            [Parameter(Mandatory)]$Export,
            [Parameter(Mandatory)]$Connection,
            [Parameter(Mandatory)]$TargetAssociated
        )

        $targetGroups = @(Get-PnPGroup -Connection $Connection)
        $targetAssociatedIds = @($TargetAssociated.Values | Where-Object { $null -ne $_ } | ForEach-Object { [int]$_.Id })
        $createdGroups = [System.Collections.Generic.List[object]]::new()

        # Passe 1 : creation / correspondance des groupes.
        foreach ($sourceGroup in @($Export.SharePointGroups)) {
            $sourceKey = 'SPGROUP:' + [string]$sourceGroup.Title

            if ($script:PrincipalMappings.ContainsKey($sourceKey)) {
                $script:TargetGroupMap[$sourceKey] = $script:PrincipalMappings[$sourceKey]
                continue
            }

            $targetGroup = $targetGroups | Where-Object Title -eq $sourceGroup.Title | Select-Object -First 1
            if ($null -ne $targetGroup -and ([int]$targetGroup.Id -in $targetAssociatedIds)) {
                Add-MigrationLog -Stage 'Groups' -Action 'Conflit avec un groupe natif cible' -Status 'Error' -Object $sourceGroup.Title -Detail 'Groupe ignore pour ne pas elargir les droits du groupe natif. Utiliser une correspondance CSV.'
                continue
            }

            if ($null -eq $targetGroup) {
                if ($Apply) {
                    try {
                        $targetGroup = New-PnPGroup -Title $sourceGroup.Title -Description $sourceGroup.Description -Connection $Connection
                        Add-MigrationLog -Stage 'Groups' -Action 'Groupe SharePoint cree' -Status 'Success' -Object $sourceGroup.Title
                    }
                    catch {
                        Add-MigrationLog -Stage 'Groups' -Action 'Creation groupe' -Status 'Error' -Object $sourceGroup.Title -Detail $_.Exception.Message
                        continue
                    }
                }
                else {
                    Add-MigrationLog -Stage 'Groups' -Action 'Creer groupe SharePoint' -Status 'Planned' -Object $sourceGroup.Title
                }
                $createdGroups.Add($sourceGroup)
            }
            else {
                Add-MigrationLog -Stage 'Groups' -Action 'Groupe existant reutilise' -Status 'Success' -Object $sourceGroup.Title
            }

            $script:TargetGroupMap[$sourceKey] = [string]$sourceGroup.Title
        }

        # Passe 2 : parametres et proprietaire des groupes crees (le proprietaire
        # peut etre un autre groupe personnalise, d'ou le traitement apres creation).
        foreach ($sourceGroup in $createdGroups) {
            $context = "Proprietaire du groupe : $($sourceGroup.Title)"
            $ownerLogin = $null
            if ($null -ne $sourceGroup.Owner) {
                $ownerLogin = Resolve-GroupOwner -Owner $sourceGroup.Owner -Connection $Connection -TargetAssociated $TargetAssociated -Context $context
            }

            if (-not $Apply) {
                Add-MigrationLog -Stage 'Groups' -Action 'Appliquer parametres/proprietaire' -Status 'Planned' -Object $sourceGroup.Title -Detail "Proprietaire: $ownerLogin"
                continue
            }

            try {
                $parameters = @{
                    Identity                       = $sourceGroup.Title
                    AllowMembersEditMembership     = [bool]$sourceGroup.AllowMembersEditMembership
                    OnlyAllowMembersViewMembership = [bool]$sourceGroup.OnlyAllowMembersViewMembership
                    AllowRequestToJoinLeave        = [bool]$sourceGroup.AllowRequestToJoinLeave
                    AutoAcceptRequestToJoinLeave   = [bool]$sourceGroup.AutoAcceptRequestToJoinLeave
                    Connection                     = $Connection
                }
                if (-not [string]::IsNullOrWhiteSpace([string]$sourceGroup.RequestToJoinLeaveEmailSetting)) {
                    $parameters.RequestToJoinEmail = [string]$sourceGroup.RequestToJoinLeaveEmailSetting
                }
                if (-not [string]::IsNullOrWhiteSpace($ownerLogin)) {
                    $parameters.Owner = $ownerLogin
                }
                Set-PnPGroup @parameters | Out-Null
                Add-MigrationLog -Stage 'Groups' -Action 'Parametres/proprietaire appliques' -Status 'Success' -Object $sourceGroup.Title -Detail "Proprietaire: $ownerLogin"
            }
            catch {
                Add-MigrationLog -Stage 'Groups' -Action 'Parametres du groupe' -Status 'Error' -Object $sourceGroup.Title -Detail $_.Exception.Message
            }
        }

        # Passe 3 : membres.
        foreach ($sourceGroup in @($Export.SharePointGroups)) {
            $sourceKey = 'SPGROUP:' + [string]$sourceGroup.Title
            if (-not $script:TargetGroupMap.ContainsKey($sourceKey)) {
                continue
            }
            $targetGroupTitle = $script:TargetGroupMap[$sourceKey]

            foreach ($member in @($sourceGroup.Members)) {
                $context = "Groupe SharePoint : $($sourceGroup.Title)"
                $resolved = Resolve-AssignmentPrincipal -Principal $member -Connection $Connection -Context $context -TargetAssociated $TargetAssociated
                if ($null -eq $resolved -or $resolved.Kind -ne 'User') {
                    continue
                }
                if ($Apply) {
                    try {
                        Add-PnPGroupMember -Group $targetGroupTitle -LoginName $resolved.Value -Connection $Connection
                        Add-MigrationLog -Stage 'Groups' -Action 'Membre ajoute' -Status 'Success' -Object $targetGroupTitle -Detail $resolved.Value
                    }
                    catch {
                        Add-MigrationLog -Stage 'Groups' -Action 'Ajout membre' -Status 'Error' -Object $targetGroupTitle -Detail "$($resolved.Value) : $($_.Exception.Message)"
                    }
                }
                else {
                    Add-MigrationLog -Stage 'Groups' -Action 'Ajouter membre' -Status 'Planned' -Object $targetGroupTitle -Detail $resolved.Value
                }
            }
        }
    }

    function Get-TargetListMap {
        param(
            [Parameter(Mandatory)]$Connection,
            [Parameter(Mandatory)][string]$TargetWebRelativeUrl
        )

        $map = @{}
        foreach ($list in Get-PnPList -Includes RootFolder, Hidden, IsSystemList -Connection $Connection) {
            $relative = ([string]$list.RootFolder.ServerRelativeUrl).Substring($TargetWebRelativeUrl.Length).TrimStart('/')
            $map[$relative.ToLowerInvariant()] = $list
        }
        return $map
    }

    function Find-TargetItemId {
        param(
            [Parameter(Mandatory)]$SourceObject,
            [Parameter(Mandatory)]$TargetList,
            [Parameter(Mandatory)][string]$SourceWebRelativeUrl,
            [Parameter(Mandatory)][string]$TargetWebRelativeUrl,
            [Parameter(Mandatory)]$Connection
        )

        $fileRef = [string]$SourceObject.FileRef
        $targetFileRef = $fileRef
        if ($fileRef.StartsWith($SourceWebRelativeUrl, [StringComparison]::OrdinalIgnoreCase)) {
            $targetFileRef = $TargetWebRelativeUrl.TrimEnd('/') + '/' + $fileRef.Substring($SourceWebRelativeUrl.Length).TrimStart('/')
        }

        try {
            if ([bool]$SourceObject.IsFolder -and -not [string]::IsNullOrWhiteSpace($fileRef)) {
                $folder = Get-PnPFolder -Url $targetFileRef -Connection $Connection
                Get-PnPProperty -ClientObject $folder -Property ListItemAllFields -Connection $Connection | Out-Null
                return [int]$folder.ListItemAllFields.Id
            }
            if ([bool]$SourceObject.IsLibrary -and -not [string]::IsNullOrWhiteSpace($fileRef)) {
                $fileItem = Get-PnPFile -Url $targetFileRef -AsListItem -Connection $Connection
                return [int]$fileItem.Id
            }
            # Liste classique : on suppose que l'outil de migration a conserve l'ID.
            $item = Get-PnPListItem -List $TargetList -Id ([int]$SourceObject.ItemId) -Connection $Connection
            return [int]$item.Id
        }
        catch {
            return $null
        }
    }

    function Clear-TargetUniquePermissions {
        param(
            [Parameter(Mandatory)]$SecurableObject,
            [Parameter(Mandatory)]$Connection,
            [Parameter(Mandatory)][string]$Context
        )

        # Miroir ou remplacement : la cible repart d'une liste vide puis recoit exactement
        # les droits source (groupes natifs transposes vers ceux du site cible).
        if (-not ($ReplaceUniquePermissions -or $MirrorSourcePermissions)) {
            return
        }

        Get-PnPProperty -ClientObject $SecurableObject -Property HasUniqueRoleAssignments -Connection $Connection | Out-Null
        $wasUnique = [bool]$SecurableObject.HasUniqueRoleAssignments
        $label = if ($wasUnique) { 'Droits cibles existants remplaces par ceux de la source' } else { 'Heritage rompu SANS copie des droits parents (miroir de la source)' }

        if (-not $Apply) {
            Add-MigrationLog -Stage 'Permissions' -Action $label -Status 'Planned' -Object $Context
            return
        }

        if ($wasUnique) {
            $SecurableObject.ResetRoleInheritance()
            Invoke-PnPQuery -Connection $Connection
        }
        $SecurableObject.BreakRoleInheritance($false, $false)
        Invoke-PnPQuery -Connection $Connection
        Add-MigrationLog -Stage 'Permissions' -Action $label -Status 'Success' -Object $Context
    }

    function Ensure-TargetUniquePermissions {
        param(
            [Parameter(Mandatory)]$SecurableObject,
            [Parameter(Mandatory)]$Connection,
            [Parameter(Mandatory)][string]$Context
        )

        # En miroir/remplacement, Clear-TargetUniquePermissions gere l'heritage.
        if ($ReplaceUniquePermissions -or $MirrorSourcePermissions) {
            return
        }

        Get-PnPProperty -ClientObject $SecurableObject -Property HasUniqueRoleAssignments -Connection $Connection | Out-Null
        if ([bool]$SecurableObject.HasUniqueRoleAssignments) {
            return
        }

        # Miroir : on part d'une liste vide pour obtenir exactement les droits source.
        $copyParent = -not $MirrorSourcePermissions
        $label = if ($copyParent) { 'avec copie des droits parents' } else { 'SANS copie des droits parents (miroir de la source)' }

        if ($Apply) {
            $SecurableObject.BreakRoleInheritance($copyParent, $false)
            Invoke-PnPQuery -Connection $Connection
            Add-MigrationLog -Stage 'Permissions' -Action "Heritage rompu $label" -Status 'Success' -Object $Context
        }
        else {
            Add-MigrationLog -Stage 'Permissions' -Action "Rompre heritage $label" -Status 'Planned' -Object $Context
        }
    }

    function Add-TargetPermission {
        param(
            [Parameter(Mandatory)][ValidateSet('Web', 'List', 'Item')][string]$ObjectType,
            $TargetList,
            [int]$TargetItemId,
            [Parameter(Mandatory)]$ResolvedPrincipal,
            [Parameter(Mandatory)][string]$RoleName,
            [Parameter(Mandatory)]$Connection,
            [Parameter(Mandatory)][string]$Context
        )

        if (-not $Apply) {
            Add-MigrationLog -Stage 'Permissions' -Action 'Ajouter autorisation' -Status 'Planned' -Object $Context -Detail "$($ResolvedPrincipal.Value) -> $RoleName"
            return
        }

        $common = @{ AddRole = $RoleName; Connection = $Connection }
        if ($ResolvedPrincipal.Kind -eq 'Group') {
            $common.Group = $ResolvedPrincipal.Value
        }
        else {
            $common.User = $ResolvedPrincipal.Value
        }

        switch ($ObjectType) {
            'Web' {
                Set-PnPWebPermission @common | Out-Null
            }
            'List' {
                $common.Identity = $TargetList
                Set-PnPListPermission @common | Out-Null
            }
            'Item' {
                $common.List = $TargetList
                $common.Identity = $TargetItemId
                if ($script:ItemPermissionSupportsSystemUpdate) {
                    $common.SystemUpdate = $true
                }
                Set-PnPListItemPermission @common | Out-Null
            }
        }
        Add-MigrationLog -Stage 'Permissions' -Action 'Autorisation ajoutee' -Status 'Success' -Object $Context -Detail "$($ResolvedPrincipal.Value) -> $RoleName"
    }

    function Get-SiteGroupId {
        param([Parameter(Mandatory)]$Connection)

        try {
            $site = Get-PnPSite -Includes GroupId -Connection $Connection
            $groupId = [guid]$site.GroupId
            if ($groupId -ne [guid]::Empty) {
                return $groupId.ToString().ToLowerInvariant()
            }
        }
        catch {
            Write-Log -Level Warn -Message "Lecture du groupe Microsoft 365 du site impossible : $($_.Exception.Message)"
        }
        return $null
    }

    function Get-JsonProp {
        param($Object, [Parameter(Mandatory)][string]$Name)

        # Lecture tolerante (Set-StrictMode) d'une propriete JSON facultative.
        if ($null -eq $Object) {
            return $null
        }
        $property = $Object.PSObject.Properties[$Name]
        if ($null -eq $property) {
            return $null
        }
        return $property.Value
    }

    function ConvertTo-EmailFromLogin {
        param([string]$Login)

        if ([string]::IsNullOrWhiteSpace($Login)) {
            return ''
        }
        $candidate = ($Login -split '\|')[-1]
        if ($candidate -match '^urn(%3a|:)spo(%3a|:)guest#(?<mail>[^@\s]+@[^@\s]+)$') {
            return $Matches['mail'].ToLowerInvariant()
        }
        if ($candidate -match '^(?<local>.+)#ext#@') {
            $local = $Matches['local']
            $lastUnderscore = $local.LastIndexOf('_')
            if ($lastUnderscore -gt 0) {
                return ($local.Substring(0, $lastUnderscore) + '@' + $local.Substring($lastUnderscore + 1)).ToLowerInvariant()
            }
        }
        if ($candidate -match '^[^@\s]+@[^@\s]+$') {
            return $candidate.ToLowerInvariant()
        }
        return ''
    }

    function Test-SourceEmailExternal {
        param(
            [Parameter(Mandatory)][string]$Email,
            [Parameter(Mandatory)]$Connection
        )

        # Externe = absent du tenant source ou compte invite (userType Guest).
        if ($script:ExternalEmailCache.ContainsKey($Email)) {
            return $script:ExternalEmailCache[$Email]
        }
        $isExternal = $true
        try {
            $sourceUser = Get-GraphUserByEmail -Email $Email -Connection $Connection
            if ($null -ne $sourceUser -and [string](Get-JsonProp $sourceUser 'userType') -eq 'Member') {
                $isExternal = $false
            }
        }
        catch {
            Write-Log -Level Warn -Message "  Impossible de verifier si '$Email' est externe (Graph User.Read.All ?) : considere comme externe."
        }
        $script:ExternalEmailCache[$Email] = $isExternal
        return $isExternal
    }

    function Get-SourceSharingGrants {
        param(
            [Parameter(Mandatory)]$Connection,
            [Parameter(Mandatory)][string]$ListId,
            [Parameter(Mandatory)][int]$ItemId,
            [Parameter(Mandatory)][string]$Context
        )

        # Lit les partages d'un fichier/dossier via Graph : liens (anonymes, organisation,
        # personnes specifiques) et invitations directes, y compris celles pas encore acceptees.
        $grants = [System.Collections.Generic.List[object]]::new()
        if ([string]::IsNullOrWhiteSpace($script:SourceGraphSiteId)) {
            return @($grants)
        }

        try {
            $response = Invoke-PnPGraphMethod -Method Get -Url "sites/$($script:SourceGraphSiteId)/lists/$ListId/items/$ItemId/driveItem/permissions" -Connection $Connection
        }
        catch {
            Add-MigrationLog -Stage 'Export' -Action 'Lecture des partages (Graph)' -Status 'Error' -Object $Context -Detail $_.Exception.Message
            return @($grants)
        }

        $allPermissions = @(Get-JsonProp $response 'value' | Where-Object { $null -ne $_ })
        Write-Log -Level Detail -Message "        Graph : $($allPermissions.Count) permission(s) lue(s) sur cet objet"
        foreach ($permission in $allPermissions) {
            # Diagnostic : resume brut de chaque permission renvoyee par Graph.
            $diagLink = Get-JsonProp $permission 'link'
            $diagInvite = Get-JsonProp $permission 'invitation'
            $diagWho = @()
            foreach ($propertyName in 'grantedToV2', 'grantedToIdentitiesV2') {
                foreach ($identity in @(Get-JsonProp $permission $propertyName)) {
                    foreach ($kind in 'user', 'siteUser', 'group', 'siteGroup') {
                        $who = Get-JsonProp $identity $kind
                        if ($null -ne $who) {
                            $whoMail = [string](Get-JsonProp $who 'email')
                            $whoName = [string](Get-JsonProp $who 'displayName')
                            $diagWho += "$kind=$(if ($whoMail) { $whoMail } else { $whoName })"
                        }
                    }
                }
            }
            Write-Log -Level Detail -Message ("          - roles={0} lien={1} invitation={2} herite={3} pour={4}" -f `
                ((@(Get-JsonProp $permission 'roles')) -join '/'),
                $(if ($diagLink) { "$(Get-JsonProp $diagLink 'scope')/$(Get-JsonProp $diagLink 'type')" } else { 'non' }),
                $(if ($diagInvite) { [string](Get-JsonProp $diagInvite 'email') } else { 'non' }),
                $(if ($null -ne (Get-JsonProp $permission 'inheritedFrom')) { 'oui' } else { 'non' }),
                $(if ($diagWho) { $diagWho -join ', ' } else { '-' }))

            if ($null -ne (Get-JsonProp $permission 'inheritedFrom')) {
                continue
            }
            $roles = @(Get-JsonProp $permission 'roles')
            if ($roles -contains 'owner') {
                continue
            }
            $link = Get-JsonProp $permission 'link'
            $invitation = Get-JsonProp $permission 'invitation'
            if ($null -eq $link -and $null -eq $invitation) {
                # Droit direct deja accorde : deja present dans les attributions SharePoint.
                continue
            }

            $scope = if ($null -ne $link) { [string](Get-JsonProp $link 'scope') } else { 'invitation' }
            $linkType = if ($null -ne $link) { [string](Get-JsonProp $link 'type') } else { 'direct' }
            $roleKind = if ($roles -contains 'write') { 'Contributor' }
                elseif ($roles -contains 'restrictedView') { 'RestrictedReader' }
                else { 'Reader' }

            # Destinataires : invitation + identites accordees (V2 puis ancienne propriete).
            $emails = [System.Collections.Generic.List[object]]::new()
            $identities = @()
            if ($null -ne $invitation) {
                $inviteMail = [string](Get-JsonProp $invitation 'email')
                if ($inviteMail) { $identities += [pscustomobject]@{ Email = $inviteMail; DisplayName = $inviteMail } }
            }
            foreach ($propertyName in 'grantedToIdentitiesV2', 'grantedToIdentities', 'grantedToV2') {
                foreach ($identity in @(Get-JsonProp $permission $propertyName)) {
                    foreach ($kind in 'user', 'siteUser') {
                        $principal = Get-JsonProp $identity $kind
                        if ($null -eq $principal) { continue }
                        $mail = [string](Get-JsonProp $principal 'email')
                        if (-not $mail) { $mail = ConvertTo-EmailFromLogin -Login ([string](Get-JsonProp $principal 'loginName')) }
                        if ($mail) {
                            $identities += [pscustomobject]@{ Email = $mail; DisplayName = [string](Get-JsonProp $principal 'displayName') }
                            break
                        }
                    }
                }
            }
            foreach ($identity in $identities) {
                $mail = ([string]$identity.Email).Trim().ToLowerInvariant()
                if (-not $mail -or @($emails | Where-Object { $_.Email -eq $mail }).Count -gt 0) { continue }
                $emails.Add([pscustomobject]@{
                    Email       = $mail
                    DisplayName = $(if ($identity.DisplayName) { $identity.DisplayName } else { $mail })
                    IsExternal  = (Test-SourceEmailExternal -Email $mail -Connection $Connection)
                })
            }

            $grants.Add([pscustomobject]@{
                Scope      = $scope
                LinkType   = $linkType
                RoleKind   = $roleKind
                Recipients = @($emails)
            })
            $recipientText = (@($emails) | ForEach-Object { "$($_.Email)$(if ($_.IsExternal) { ' [EXTERNE]' })" }) -join ', '
            Write-Log -Level Info -Message "        Partage trouve : $scope/$linkType ($roleKind) -> $(if ($recipientText) { $recipientText } else { '(aucun destinataire nominatif)' })"
        }
        return @($grants)
    }

    function Expand-SharingLinkAssignments {
        param(
            [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Assignments,
            [Parameter(Mandatory)]$SharingLinks,
            [Parameter(Mandatory)][string]$Context,
            [AllowEmptyCollection()][object[]]$SharingGrants = @()
        )

        # Remplace chaque attribution "lien de partage" par des attributions directes
        # a ses destinataires (memes niveaux d'autorisation). Aucun e-mail n'est envoye.
        $result = [System.Collections.Generic.List[object]]::new()
        foreach ($assignment in $Assignments) {
            if ([string]$assignment.Principal.Category -ne 'SharingLink') {
                $result.Add($assignment)
                continue
            }

            $linkTitle = [string]$assignment.Principal.Title
            if ($SharingLinkRecipients -eq 'None' -or -not $SharingLinks.ContainsKey($linkTitle)) {
                Add-MigrationLog -Stage 'Partage' -Action 'Lien de partage ignore' -Status 'Skipped' -Object $Context -Detail $linkTitle
                continue
            }

            $link = $SharingLinks[$linkTitle]
            $recipients = @($link.Members | Where-Object {
                [string]$_.Category -eq 'User' -and ($SharingLinkRecipients -eq 'All' -or [bool]$_.IsExternal)
            })
            $roleNames = (@($assignment.Roles) | ForEach-Object { $_.Name }) -join ', '

            if ($recipients.Count -eq 0) {
                $reason = if ([string]$link.LinkType -like 'Anonymous*') {
                    'Lien "Tout le monde" (anonyme) : aucun destinataire nominatif, non recreable.'
                }
                elseif ([string]$link.LinkType -like 'Organization*') {
                    'Lien "Personnes de l''organisation" : aucun destinataire externe.'
                }
                else {
                    "Aucun destinataire correspondant au filtre SharingLinkRecipients = $SharingLinkRecipients."
                }
                Add-MigrationLog -Stage 'Partage' -Action "Lien $($link.LinkType) sans destinataire a reporter" -Status 'Skipped' -Object $Context -Detail $reason
                continue
            }

            foreach ($recipient in $recipients) {
                Add-MigrationLog -Stage 'Partage' -Action "Partage $($link.LinkType) -> droit direct" -Status 'Planned' -Object $Context -Detail "$($recipient.Title) <$($recipient.Email)> $(if ($recipient.IsExternal) { '[EXTERNE]' } else { '[interne]' }) : $roleNames"
                $result.Add([pscustomobject]@{
                    Principal = $recipient
                    Roles     = @($assignment.Roles)
                })
            }
        }

        # Partages lus via Graph (liens personnes specifiques + invitations en attente).
        foreach ($grant in $SharingGrants) {
            if ([string]$grant.Scope -in @('anonymous', 'organization')) {
                $label = if ([string]$grant.Scope -eq 'anonymous') { 'Tout le monde (anonyme)' } else { 'Personnes de l''organisation' }
                Add-MigrationLog -Stage 'Partage' -Action "Lien '$label' non recreable" -Status 'Skipped' -Object $Context -Detail "type $($grant.LinkType) : a recreer manuellement si necessaire."
                continue
            }
            foreach ($recipient in @($grant.Recipients)) {
                if ($SharingLinkRecipients -ne 'All' -and -not [bool]$recipient.IsExternal) {
                    Write-Log -Level Detail -Message "    partage vers $($recipient.Email) ignore (interne, SharingLinkRecipients = $SharingLinkRecipients)"
                    continue
                }
                $principal = [pscustomobject]@{
                    Key           = 'USER:' + [string]$recipient.Email
                    Title         = [string]$recipient.DisplayName
                    Email         = [string]$recipient.Email
                    LoginName     = ''
                    PrincipalType = 'User'
                    Category      = 'User'
                    Association   = 'None'
                    IsExternal    = [bool]$recipient.IsExternal
                }
                $role = [pscustomobject]@{ Name = [string]$grant.RoleKind; RoleTypeKind = [string]$grant.RoleKind; IsCustom = $false }
                Add-MigrationLog -Stage 'Partage' -Action "Partage $($grant.Scope)/$($grant.LinkType) -> droit direct" -Status 'Planned' -Object $Context -Detail "$($recipient.Email) $(if ($recipient.IsExternal) { '[EXTERNE]' } else { '[interne]' }) : $($grant.RoleKind), sans e-mail"
                $result.Add([pscustomobject]@{
                    Principal = $principal
                    Roles     = @($role)
                })
            }
        }
        return @($result)
    }

    function Test-AssignmentInScope {
        param(
            [Parameter(Mandatory)]$Assignment,
            [Parameter(Mandatory)][string]$ObjectType
        )

        # Au niveau du site, les groupes natifs ont deja leurs droits natifs dans la cible :
        # on ne reporte que leurs niveaux personnalises.
        if ([string]$Assignment.Principal.Category -in @('AssociatedGroup', 'SiteM365Group') -and $ObjectType -eq 'Web') {
            return (@($Assignment.Roles | Where-Object { [bool]$_.IsCustom }).Count -gt 0)
        }
        return $true
    }

    function Import-SharePointSecurity {
        param(
            [Parameter(Mandatory)]$Export,
            [Parameter(Mandatory)]$Connection
        )

        Import-PrincipalMappings
        Write-Log -Level Detail -Message "$($script:PrincipalMappings.Count) correspondance(s) chargee(s) depuis le CSV."
        $targetWeb = Get-PnPWeb -Includes ServerRelativeUrl, Title -Connection $Connection
        $targetWebRelativeUrl = [string]$targetWeb.ServerRelativeUrl
        $sourceWebRelativeUrl = [string]$Export.SourceWebRelativeUrl
        Write-Log -Message "Site cible : '$($targetWeb.Title)' ($targetWebRelativeUrl)"
        $script:TargetSiteGroupId = Get-SiteGroupId -Connection $Connection
        Write-Log -Message "Groupe Microsoft 365 du site cible : $(if ($script:TargetSiteGroupId) { $script:TargetSiteGroupId } else { '(site non connecte a un groupe)' })"
        $targetAssociated = Get-AssociatedGroups -Connection $Connection -Stage 'Import'
        foreach ($association in $targetAssociated.Keys) {
            if ($null -ne $targetAssociated[$association]) {
                Write-Log -Level Detail -Message "  Groupe natif cible $association = '$($targetAssociated[$association].Title)'"
            }
        }

        Start-Step -Name "Niveaux d'autorisation personnalises ($(@($Export.CustomRoles).Count))"
        Initialize-TargetRoles -Export $Export -Connection $Connection

        Start-Step -Name "Groupes SharePoint personnalises et membres ($(@($Export.SharePointGroups).Count))"
        Initialize-TargetGroups -Export $Export -Connection $Connection -TargetAssociated $targetAssociated

        $objects = @($Export.SecurableObjects)
        Start-Step -Name "Autorisations sur le site, les listes et les elements ($($objects.Count) objet(s))"
        $targetLists = Get-TargetListMap -Connection $Connection -TargetWebRelativeUrl $targetWebRelativeUrl
        Write-Log -Level Detail -Message "$($targetLists.Count) liste(s) trouvee(s) sur le site cible."

        $sharingLinks = @{}
        foreach ($link in @($Export.SharingLinkGroups)) {
            $sharingLinks[[string]$link.Title] = $link
        }
        Write-Log -Message "Partages : $($sharingLinks.Count) lien(s) de partage source, destinataires reportes : $SharingLinkRecipients (droits directs, sans e-mail)."

        $objectIndex = 0
        foreach ($sourceObject in $objects) {
            $objectIndex++
            $objectType = [string]$sourceObject.ObjectType

            $context = switch ($objectType) {
                'Web'  { 'Site racine' }
                'List' { "Liste : $($sourceObject.ListRelativeUrl)" }
                default { "Element : $($sourceObject.RelativePath)" }
            }
            Show-Progress -Activity 'Import des autorisations SharePoint' -Status $context -Index $objectIndex -Total $objects.Count
            Write-Log -Level Info -Message "[$objectIndex/$($objects.Count)] $context - $(@($sourceObject.Assignments).Count) attribution(s) source"

            try {
                $targetList = $null
                $targetItemId = 0
                $targetSecurableObject = $null

                if ($objectType -in @('List', 'Item')) {
                    $listKey = ([string]$sourceObject.ListRelativeUrl).ToLowerInvariant()
                    if (-not $targetLists.ContainsKey($listKey)) {
                        Add-MigrationLog -Stage 'Permissions' -Action 'Objet cible introuvable' -Status 'Skipped' -Object $context -Detail 'Bibliotheque ou liste absente.'
                        continue
                    }
                    $targetList = $targetLists[$listKey]
                }

                if ($objectType -eq 'List') {
                    $targetSecurableObject = $targetList
                }
                elseif ($objectType -eq 'Item') {
                    $foundId = Find-TargetItemId -SourceObject $sourceObject -TargetList $targetList `
                        -SourceWebRelativeUrl $sourceWebRelativeUrl -TargetWebRelativeUrl $targetWebRelativeUrl -Connection $Connection
                    if ($null -eq $foundId) {
                        Add-MigrationLog -Stage 'Permissions' -Action 'Element cible introuvable' -Status 'Skipped' -Object $context -Detail 'Chemin ou ID non retrouve.'
                        continue
                    }
                    $targetItemId = [int]$foundId
                    $targetSecurableObject = Get-PnPListItem -List $targetList -Id $targetItemId -Connection $Connection
                    Write-Log -Level Detail -Message "  Element cible trouve : ID $targetItemId (source ID $($sourceObject.ItemId))"
                }

                $resolvedAssignments = [System.Collections.Generic.List[object]]::new()
                $sourceAssignments = Expand-SharingLinkAssignments -Assignments @($sourceObject.Assignments) -SharingLinks $sharingLinks -Context $context -SharingGrants @($sourceObject.SharingGrants)
                foreach ($assignment in $sourceAssignments) {
                    $roleList = (@($assignment.Roles) | ForEach-Object { $_.Name }) -join ', '
                    Write-Log -Level Detail -Message "  Source : $($assignment.Principal.Title) [$($assignment.Principal.Category)] -> $roleList"
                    if (-not (Test-AssignmentInScope -Assignment $assignment -ObjectType $objectType)) {
                        Write-Log -Level Detail -Message '    ignore : groupe natif avec niveaux natifs au niveau du site.'
                        continue
                    }
                    $resolvedPrincipal = Resolve-AssignmentPrincipal -Principal $assignment.Principal -Connection $Connection -Context $context -TargetAssociated $targetAssociated
                    if ($null -eq $resolvedPrincipal) {
                        continue
                    }

                    $roleNames = [System.Collections.Generic.List[string]]::new()
                    foreach ($role in @($assignment.Roles)) {
                        # Au niveau site, un groupe natif ne recoit que ses niveaux personnalises.
                        if ($objectType -eq 'Web' -and [string]$assignment.Principal.Category -in @('AssociatedGroup', 'SiteM365Group') -and -not [bool]$role.IsCustom) {
                            continue
                        }
                        $targetRoleName = Resolve-TargetRoleName -Role $role
                        if ([string]::IsNullOrWhiteSpace($targetRoleName)) {
                            Add-MigrationLog -Stage 'Permissions' -Action 'Niveau cible introuvable' -Status 'Skipped' -Object $context -Detail $role.Name
                            continue
                        }
                        $roleNames.Add($targetRoleName)
                    }

                    if ($roleNames.Count -gt 0) {
                        $resolvedAssignments.Add([pscustomobject]@{
                            Principal = $resolvedPrincipal
                            Roles     = @($roleNames)
                        })
                    }
                }

                if ($resolvedAssignments.Count -eq 0) {
                    Add-MigrationLog -Stage 'Permissions' -Action 'Aucune autorisation non native a reporter' -Status 'Skipped' -Object $context
                    continue
                }
                Write-Log -Level Detail -Message "  $($resolvedAssignments.Count) attribution(s) a appliquer sur la cible."

                if ($objectType -ne 'Web') {
                    Ensure-TargetUniquePermissions -SecurableObject $targetSecurableObject -Connection $Connection -Context $context
                    Clear-TargetUniquePermissions -SecurableObject $targetSecurableObject -Connection $Connection -Context $context
                }

                foreach ($assignment in $resolvedAssignments) {
                    foreach ($roleName in $assignment.Roles) {
                        try {
                            Add-TargetPermission -ObjectType $objectType -TargetList $targetList `
                                -TargetItemId $targetItemId -ResolvedPrincipal $assignment.Principal -RoleName $roleName `
                                -Connection $Connection -Context $context
                        }
                        catch {
                            Add-MigrationLog -Stage 'Permissions' -Action 'Ajout autorisation' -Status 'Error' -Object $context -Detail "$($assignment.Principal.Value) -> $roleName : $($_.Exception.Message)"
                        }
                    }
                }
            }
            catch {
                Add-MigrationLog -Stage 'Permissions' -Action 'Traitement objet' -Status 'Error' -Object $context -Detail $_.Exception.Message
            }
        }
        Write-Progress -Activity 'Import des autorisations SharePoint' -Completed
    }

    function Save-Reports {
        param([Parameter(Mandatory)][string]$Directory)

        $logPath = Join-Path $Directory 'MigrationLog.csv'
        $unresolvedPath = Join-Path $Directory 'UnresolvedPrincipals.csv'
        $invitationsPath = Join-Path $Directory 'Invitations.csv'

        $script:MigrationLog | Export-Csv -LiteralPath $logPath -NoTypeInformation -Encoding utf8
        $script:Unresolved | Sort-Object SourceKey, Context -Unique | Export-Csv -LiteralPath $unresolvedPath -NoTypeInformation -Encoding utf8
        $script:Invitations | Sort-Object Email -Unique | Export-Csv -LiteralPath $invitationsPath -NoTypeInformation -Encoding utf8

        $errors = @($script:MigrationLog | Where-Object Status -eq 'Error').Count
        Write-Host "`nRapports ($errors erreur(s), $($script:Unresolved.Count) principal(aux) non resolu(s)) :" -ForegroundColor Cyan
        Write-Host "- $logPath"
        Write-Host "- $unresolvedPath"
        Write-Host "- $invitationsPath"
    }

    function Assert-Configuration {
        if ($SharingLinkRecipients -notin @('External', 'All', 'None')) {
            throw "Configuration invalide pour SharingLinkRecipients : '$SharingLinkRecipients' (valeurs possibles : 'External', 'All', 'None')."
        }

        foreach ($siteSetting in @(
                @{ Name = 'SourceSiteUrl'; Value = $SourceSiteUrl },
                @{ Name = 'TargetSiteUrl'; Value = $TargetSiteUrl }
            )) {
            $parsedUri = $null
            if (-not [Uri]::TryCreate([string]$siteSetting.Value, [UriKind]::Absolute, [ref]$parsedUri) -or
                $parsedUri.Scheme -ne 'https' -or
                $parsedUri.Host -notlike '*.sharepoint.com') {
                throw "Configuration invalide pour $($siteSetting.Name) : indiquez l'URL SharePoint complete."
            }
        }

        foreach ($clientSetting in @(
                @{ Name = 'SourceClientId'; Value = $SourceClientId },
                @{ Name = 'TargetClientId'; Value = $TargetClientId }
            )) {
            if (-not (Test-IsGuid -Value ([string]$clientSetting.Value))) {
                throw "Configuration invalide pour $($clientSetting.Name) : indiquez l'ID d'application Entra correspondant."
            }
        }

        if (-not $script:SameTenant -and $SourceClientId -eq $TargetClientId) {
            throw 'Cross-tenant : SourceClientId et TargetClientId doivent correspondre a deux applications distinctes (une par tenant).'
        }

        if ($SourceSiteUrl.TrimEnd('/') -eq $TargetSiteUrl.TrimEnd('/')) {
            throw 'SourceSiteUrl et TargetSiteUrl designent le meme site.'
        }

        foreach ($tenantSetting in @(
                @{ Name = 'SourceTenantId'; Value = $SourceTenantId },
                @{ Name = 'TargetTenantId'; Value = $TargetTenantId }
            )) {
            if (-not (Test-IsGuid -Value ([string]$tenantSetting.Value)) -and
                [string]$tenantSetting.Value -notmatch '^[A-Za-z0-9.-]+\.onmicrosoft\.com$') {
                throw "Configuration invalide pour $($tenantSetting.Name) : indiquez le GUID du tenant ou son domaine tenant.onmicrosoft.com."
            }
        }


        foreach ($certificateSetting in @(
                @{ Name = 'certificat source'; Path = $SourceCertificatePath; Password = $SourceCertificatePassword },
                @{ Name = 'certificat cible'; Path = $TargetCertificatePath; Password = $TargetCertificatePassword }
            )) {
            if ([string]::IsNullOrWhiteSpace([string]$certificateSetting.Path) -or
                -not (Test-Path -LiteralPath $certificateSetting.Path -PathType Leaf)) {
                throw "Fichier introuvable pour le $($certificateSetting.Name) : $($certificateSetting.Path)"
            }
            if ([IO.Path]::GetExtension([string]$certificateSetting.Path) -notin @('.pfx', '.p12')) {
                throw "Le $($certificateSetting.Name) doit etre un fichier .pfx ou .p12."
            }
            if ([string]::IsNullOrWhiteSpace([string]$certificateSetting.Password)) {
                throw "Configuration invalide pour le $($certificateSetting.Name) : indiquez son mot de passe."
            }

            $certificate = $null
            try {
                $resolvedCertificatePath = (Resolve-Path -LiteralPath $certificateSetting.Path).Path
                # EphemeralKeySet n'est pas fiable sous .NET Framework (Windows PowerShell 5.1).
                $keyFlags = if ($PSVersionTable.PSVersion.Major -ge 7) {
                    [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
                }
                else {
                    [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::DefaultKeySet
                }
                $certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(
                    $resolvedCertificatePath,
                    [string]$certificateSetting.Password,
                    $keyFlags
                )
                if (-not $certificate.HasPrivateKey) {
                    throw 'Le fichier ne contient pas de cle privee.'
                }
                if ($certificate.NotAfter -le (Get-Date)) {
                    throw "Le certificat a expire le $($certificate.NotAfter)."
                }
                if ($certificate.NotBefore -gt (Get-Date)) {
                    throw "Le certificat ne sera valide qu'a partir du $($certificate.NotBefore)."
                }
            }
            catch {
                throw "Impossible d'ouvrir le $($certificateSetting.Name). Verifiez le chemin, le mot de passe et la cle privee : $($_.Exception.Message)"
            }
            finally {
                if ($null -ne $certificate) {
                    $certificate.Dispose()
                }
            }
        }

        if (-not $script:SameTenant -and
            (Resolve-Path -LiteralPath $SourceCertificatePath).Path -eq (Resolve-Path -LiteralPath $TargetCertificatePath).Path) {
            throw 'Cross-tenant : les certificats source et cible doivent etre deux fichiers distincts.'
        }
    }

    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    $runStamp = '{0:yyyyMMdd_HHmmss}' -f (Get-Date)
    $exportFile = Join-Path $OutputDirectory 'SharePointPermissions.json'
    $transcriptFile = Join-Path $OutputDirectory "Execution_$runStamp.log"
    $script:LiveLogFile = Join-Path $OutputDirectory "Suivi_$runStamp.log"
    $sourceConnection = $null
    $targetConnection = $null
    $transcriptStarted = $false

    try {
        Start-Transcript -LiteralPath $transcriptFile -Force | Out-Null
        $transcriptStarted = $true
    }
    catch {
        Write-Warning "Impossible de demarrer le journal d'execution : $($_.Exception.Message)"
    }

    Write-Log -Level Step -Message ("=" * 70)
    Write-Log -Level Step -Message 'MIGRATION DES DROITS SHAREPOINT NON NATIFS (meme tenant ou cross-tenant)'
    Write-Log -Level Step -Message ("=" * 70)
    Write-Log -Message "PowerShell $($PSVersionTable.PSVersion) - PnP.PowerShell $($pnpModule.Version) - hote : $($Host.Name)"
    Write-Log -Message "Script    : $PSCommandPath"
    Write-Log -Message "Rapports  : $OutputDirectory"
    Write-Log -Message "Suivi     : $($script:LiveLogFile)"
    Write-Log -Message "Mode      : $(if ($Apply) { 'APPLICATION (modifications reelles)' } else { 'SIMULATION (aucune modification)' })"
    Write-Log -Message "Options   : MirrorSourcePermissions=$MirrorSourcePermissions, ReplaceUniquePermissions=$ReplaceUniquePermissions, IncludeHiddenLists=$IncludeHiddenLists, SkipItemPermissions=$SkipItemPermissions, InviteMissingUsers=$InviteMissingUsers, IncludeAssociatedGroupAssignments=$IncludeAssociatedGroupAssignments, MigrateEveryoneClaims=$MigrateEveryoneClaims, SharingLinkRecipients=$SharingLinkRecipients, InviteExternalRecipients=$InviteExternalRecipients, LogLevel=$LogLevel"
    Write-Log -Message "Domaines  : $(if ($DomainMapping.Count -gt 0) { ($DomainMapping.GetEnumerator() | ForEach-Object { "$($_.Key) -> $($_.Value)" }) -join ', ' } else { '(aucune correspondance)' })"

    if (Test-IsGuid -Value $TargetTenantId) {
        $script:TargetTenantGuid = $TargetTenantId.ToLowerInvariant()
    }

    $script:SameTenant = ([string]$SourceTenantId).Trim().ToLowerInvariant() -eq ([string]$TargetTenantId).Trim().ToLowerInvariant()
    if ($script:SameTenant) {
        Write-Log -Level Warn -Message 'MEME TENANT detecte (source et cible identiques) : copie intra-tenant, les comptes et groupes Entra sont reutilises tels quels.'
        Write-Log -Level Warn -Message 'La meme application / le meme certificat peuvent servir aux deux connexions (droits FullControl requis sur les deux sites).'
    }
    else {
        Write-Log -Message 'Mode CROSS-TENANT.'
    }

    try {
        Start-Step -Name 'Verification de la configuration et connexion au site source'
        Write-Log -Message "Source : $SourceSiteUrl (tenant $SourceTenantId, application $SourceClientId)"
        Write-Log -Message "Cible  : $TargetSiteUrl (tenant $TargetTenantId, application $TargetClientId)"
        Assert-Configuration
        Write-Log -Level Success -Message 'Configuration valide (URL, identifiants, certificats).'

        Write-Log -Message 'Connexion applicative au site source avec le certificat source...'
        $sourceConnection = Connect-PnPOnline `
            -Url $SourceSiteUrl `
            -Tenant $SourceTenantId `
            -ClientId $SourceClientId `
            -CertificatePath $SourceCertificatePath `
            -CertificatePassword (ConvertTo-SecureString -String $SourceCertificatePassword -AsPlainText -Force) `
            -ReturnConnection
        Write-Log -Level Success -Message 'Connecte au site source.'

        Start-Step -Name 'Export des droits non natifs du site source'
        $export = Export-SharePointSecurity -Connection $sourceConnection -ExportFile $exportFile

        Write-Log -Message 'Connexion applicative au site cible avec le certificat cible...'
        $targetConnection = Connect-PnPOnline `
            -Url $TargetSiteUrl `
            -Tenant $TargetTenantId `
            -ClientId $TargetClientId `
            -CertificatePath $TargetCertificatePath `
            -CertificatePassword (ConvertTo-SecureString -String $TargetCertificatePassword -AsPlainText -Force) `
            -ReturnConnection
        Write-Log -Level Success -Message 'Connecte au site cible.'

        if (-not $Apply) {
            Write-Log -Level Warn -Message 'MODE SIMULATION : aucune modification ne sera effectuee.'
        }
        elseif ($ReplaceUniquePermissions -or $MirrorSourcePermissions) {
            Write-Log -Level Warn -Message 'MODE APPLICATION (miroir) : les droits des objets a droits uniques seront remplaces par ceux de la source.'
        }
        else {
            Write-Log -Level Warn -Message 'MODE APPLICATION : les autorisations seront ajoutees/fusionnees.'
        }

        Import-SharePointSecurity -Export $export -Connection $targetConnection
    }
    catch {
        $failure = $_
        Add-MigrationLog -Stage 'General' -Action 'Execution interrompue' -Status 'Error' -Object '' -Detail $failure.Exception.Message
        Write-Log -Level Error -Message '==================== ERREUR ===================='
        Write-Log -Level Error -Message $failure.Exception.Message
        Write-Log -Level Error -Message "Ligne $($failure.InvocationInfo.ScriptLineNumber) : $($failure.InvocationInfo.Line.Trim())"
        Write-Log -Level Error -Message '================================================'
        $script:PhaseFailed = $true
    }
    finally {
        Write-ExecutionSummary
        Save-Reports -Directory $OutputDirectory
        $sourceConnection = $null
        $targetConnection = $null
        Write-Log -Message "Journal de suivi : $($script:LiveLogFile)"
        if ($transcriptStarted) {
            Stop-Transcript | Out-Null
            Write-Host "Journal complet : $transcriptFile" -ForegroundColor Cyan
        }
    }
}

function Invoke-PhaseMetadonnees {
    # Phase Metadonnees : remet Cree/Modifie (dates) et Cree par/Modifie par d'origine sur la cible (PnP, SystemUpdate)
    $OutputDirectory = Join-Path $WorkDirectory 'SharePoint-Metadata'
    $ErrorActionPreference = 'Stop'
    $script:LogFile = $null
    $results = New-Object System.Collections.Generic.List[object]
    $stats = @{ ok = 0; datesOnly = 0; failed = 0; skipped = 0; mismatch = 0 }
    $userIdCache = @{}

    function Write-Log {
        param([string]$Message, [string]$Level = 'INFO')
        $line = "[{0}][{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
        $color = switch ($Level) { 'ERROR' { 'Red' } 'WARN' { 'Yellow' } 'SUCCESS' { 'Green' } default { 'White' } }
        Write-Host $line -ForegroundColor $color
        if ($script:LogFile) { try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -ErrorAction Stop } catch {} }
    }

    function Read-Jsonl {
        param([string]$Path)
        if (-not (Test-Path -LiteralPath $Path)) { return }
        foreach ($line in [System.IO.File]::ReadLines($Path, [System.Text.Encoding]::UTF8)) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $line | ConvertFrom-Json
        }
    }

    function Connect-Site {
        param([string]$Url, [string]$Tenant, [string]$ClientId, [string]$CertPath, [string]$CertPassword)
        if (-not (Test-Path -LiteralPath $CertPath -PathType Leaf)) { throw "Certificat introuvable: $CertPath" }
        if ([string]::IsNullOrWhiteSpace($CertPassword)) { throw "Mot de passe du certificat vide pour $CertPath" }
        return Connect-PnPOnline -Url $Url -Tenant $Tenant -ClientId $ClientId -CertificatePath $CertPath `
            -CertificatePassword (ConvertTo-SecureString -String $CertPassword -AsPlainText -Force) -ReturnConnection
    }

    # Adresse e-mail source -> adresse cible (correspondance de domaines de $DomainMapping, comme la phase Droits).
    function Convert-Email([string]$Email) {
        if ([string]::IsNullOrWhiteSpace($Email)) { return $null }
        $parts = $Email.Trim().Split('@')
        if ($parts.Count -eq 2 -and $DomainMapping.ContainsKey($parts[1])) {
            return ($parts[0] + '@' + [string]$DomainMapping[$parts[1]]).ToLowerInvariant()
        }
        return $Email.Trim().ToLowerInvariant()
    }

    function ConvertTo-UtcDate($Value) {
        if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
        return [datetime]::Parse([string]$Value, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()
    }

    # Adresse -> id de l'utilisateur dans la liste d'information utilisateurs du site cible
    # (l'adresse e-mail est d'abord convertie en UPN via Graph, comme dans la phase Droits).
    function Resolve-TargetUserId {
        param([string]$Email, $Conn)

        $mail = Convert-Email $Email
        if (-not $mail) { return $null }
        if ($userIdCache.ContainsKey($mail)) { return $userIdCache[$mail] }

        $id = $null
        $upn = $mail
        try {
            $filter = [Uri]::EscapeDataString("mail eq '$mail' or userPrincipalName eq '$mail'")
            $resp = Invoke-PnPGraphMethod -Method Get -Url "users?`$select=userPrincipalName&`$filter=$filter" -Connection $Conn
            $u = @($resp.value) | Select-Object -First 1
            if ($u -and $u.userPrincipalName) { $upn = ([string]$u.userPrincipalName).ToLowerInvariant() }
        } catch {
            Write-Log ("Recherche Graph impossible pour {0} : {1}" -f $mail, $_.Exception.Message) 'WARN'
        }
        try {
            $siteUser = New-PnPUser -LoginName ('i:0#.f|membership|' + $upn) -Connection $Conn
            $id = [int]$siteUser.Id
        } catch {
            Write-Log ("Utilisateur introuvable sur la cible : {0} ({1})" -f $upn, $_.Exception.Message) 'WARN'
        }
        $userIdCache[$mail] = $id
        return $id
    }

    function New-UserValue([int]$Id) {
        $v = New-Object Microsoft.SharePoint.Client.FieldUserValue
        $v.LookupId = $Id
        return $v
    }

    # Ecrit Cree/Modifie/Cree par/Modifie par (SystemUpdate : ne change ni la date ni l'auteur de modification),
    # puis relit l'element pour verifier ce que SharePoint a reellement enregistre.
    function Set-Metadata {
        param($List, [int]$ItemId, $Meta, $Conn, [string]$Label)

        $created  = if ($MetadataSetDates -and $Meta.created)  { ConvertTo-UtcDate $Meta.created }  else { $null }
        $modified = if ($MetadataSetDates -and $Meta.modified) { ConvertTo-UtcDate $Meta.modified } else { $null }
        $authorMail = if ($MetadataSetAuthors -and $Meta.createdBy)  { [string]$Meta.createdBy.email }  else { '' }
        $editorMail = if ($MetadataSetAuthors -and $Meta.modifiedBy) { [string]$Meta.modifiedBy.email } else { '' }
        if (-not $created -and -not $modified -and -not $authorMail -and -not $editorMail) { $stats.skipped++; return 'ignore' }
        if (-not $Apply) { $stats.ok++; return 'prevu' }

        try {
            $authorId = if ($authorMail) { Resolve-TargetUserId -Email $authorMail -Conn $Conn } else { $null }
            $editorId = if ($editorMail) { Resolve-TargetUserId -Email $editorMail -Conn $Conn } else { $null }

            $item = Get-PnPListItem -List $List -Id $ItemId -Connection $Conn
            if ($created)  { $item['Created']  = $created }
            if ($modified) { $item['Modified'] = $modified }
            if ($authorId) { $item['Author'] = New-UserValue $authorId }
            if ($editorId) { $item['Editor'] = New-UserValue $editorId }
            $item.SystemUpdate()
            Invoke-PnPQuery -Connection $Conn
        } catch {
            $stats.failed++
            Write-Log ("{0} : ecriture impossible : {1}" -f $Label, $_.Exception.Message) 'ERROR'
            return ('erreur: ' + $_.Exception.Message)
        }

        $notes = New-Object System.Collections.Generic.List[string]
        if ($authorMail -and -not $authorId) { $notes.Add("auteur '$authorMail' introuvable sur la cible") }
        if ($editorMail -and -not $editorId) { $notes.Add("modificateur '$editorMail' introuvable sur la cible") }

        if ($MetadataVerify) {
            try {
                $check = Get-PnPListItem -List $List -Id $ItemId -Connection $Conn
                $diffs = New-Object System.Collections.Generic.List[string]
                foreach ($pair in @(@('Created', $created), @('Modified', $modified))) {
                    if ($null -eq $pair[1]) { continue }
                    $actual = [datetime]$check[$pair[0]]
                    if ([Math]::Abs(($actual.ToUniversalTime() - $pair[1]).TotalSeconds) -gt 2) {
                        $diffs.Add(("{0} attendu {1:u} / obtenu {2:u}" -f $pair[0], $pair[1], $actual.ToUniversalTime()))
                    }
                }
                foreach ($pair in @(@('Author', $authorId), @('Editor', $editorId))) {
                    if (-not $pair[1]) { continue }
                    $actualId = [int]$check[$pair[0]].LookupId
                    if ($actualId -ne [int]$pair[1]) {
                        $diffs.Add(("{0} attendu id {1} / obtenu '{2}' (id {3})" -f $pair[0], $pair[1], $check[$pair[0]].LookupValue, $actualId))
                    }
                }
                if ($diffs.Count -gt 0) {
                    $stats.mismatch++
                    Write-Log ("{0} : valeurs non retenues : {1}" -f $Label, ($diffs -join ' ; ')) 'WARN'
                    return ('ecart: ' + ($diffs -join ' ; '))
                }
            } catch {
                Write-Log ("{0} : verification impossible : {1}" -f $Label, $_.Exception.Message) 'WARN'
            }
        }

        if ($notes.Count -gt 0) {
            $stats.datesOnly++
            Write-Log ("{0} : {1}" -f $Label, ($notes -join ' ; ')) 'WARN'
            return ('partiel: ' + ($notes -join ' ; '))
        }
        $stats.ok++
        return 'ok'
    }

    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    $script:LogFile = Join-Path $OutputDirectory 'metadonnees.log'
    Set-Content -LiteralPath $script:LogFile -Value '' -Encoding UTF8
    Write-Log ("Mode: {0}" -f $(if ($Apply) { 'APPLICATION' } else { 'SIMULATION (aucune modification)' }))

    try {
        $dstConn = $null
        $dstWeb = $null
        $listByRel = @{}
        if ($Apply) {
            $dstConn = Connect-Site -Url $TargetSiteUrl -Tenant $TargetTenantId -ClientId $TargetClientId -CertPath $TargetCertificatePath -CertPassword $TargetCertificatePassword
            $dstWeb = (Get-PnPWeb -Connection $dstConn).ServerRelativeUrl.TrimEnd('/')
            foreach ($l in @(Get-PnPList -Includes RootFolder -Connection $dstConn)) {
                $rel = ([string]$l.RootFolder.ServerRelativeUrl).Substring($dstWeb.Length).TrimStart('/').ToLowerInvariant()
                $listByRel[$rel] = $l
            }
        }

        # ---- Fichiers et dossiers des bibliotheques ----
        $copyResults = Join-Path $WorkDirectory 'SharePoint-Copy\resultats_fichiers.csv'
        $itemsFile = Join-Path $DiscoveryDirectory 'items.jsonl'
        if (-not (Test-Path -LiteralPath $copyResults)) { throw "Resultats de copie introuvables : $copyResults (lancez d'abord la phase Copie)" }
        if (-not (Test-Path -LiteralPath $itemsFile)) { throw "items.jsonl introuvable dans $DiscoveryDirectory (lancez la phase Decouverte)" }

        $metaByPath = @{}
        $withMeta = 0
        foreach ($row in (Read-Jsonl -Path $itemsFile)) {
            if ($row.PSObject.Properties['meta'] -and $row.meta) { $metaByPath[[string]$row.path] = $row.meta; $withMeta++ }
        }
        if ($withMeta -eq 0) {
            Write-Log 'Aucune metadonnee dans items.jsonl : relancez la phase Decouverte (version recente) puis la Copie.' 'WARN'
        }

        $copied = @(Import-Csv -LiteralPath $copyResults -Encoding UTF8 | Where-Object { $_.Status -eq 'copied' -and $_.PathTranslated })
        Write-Log ("Fichiers/dossiers copies a traiter : {0}" -f $copied.Count)
        $i = 0
        foreach ($r in $copied) {
            $i++
            $meta = $metaByPath[[string]$r.PathOriginal]
            if (-not $meta) { $stats.skipped++; continue }
            $translated = ([string]$r.PathTranslated).Trim('/')
            $libKey = ($translated -split '/')[0]
            $label = $translated
            if (-not $Apply) { $null = Set-Metadata -List $null -ItemId 0 -Meta $meta -Conn $null -Label $label; continue }

            try {
                $list = $listByRel[$libKey.ToLowerInvariant()]
                if (-not $list) { throw "bibliotheque cible introuvable : $libKey" }
                $serverRel = $dstWeb + '/' + $translated
                if ([string]$r.ItemType -eq 'Folder') {
                    $folder = Get-PnPFolder -Url $serverRel -Connection $dstConn
                    Get-PnPProperty -ClientObject $folder -Property ListItemAllFields -Connection $dstConn | Out-Null
                    $id = [int]$folder.ListItemAllFields.Id
                } else {
                    $id = [int](Get-PnPFile -Url $serverRel -AsListItem -Connection $dstConn).Id
                }
                $res = Set-Metadata -List $list -ItemId $id -Meta $meta -Conn $dstConn -Label $label
                $results.Add([pscustomobject]@{ Type = 'Fichier'; Chemin = $translated; Resultat = $res }) | Out-Null
            } catch {
                $stats.failed++
                Write-Log ("{0} : {1}" -f $label, $_.Exception.Message) 'ERROR'
                $results.Add([pscustomobject]@{ Type = 'Fichier'; Chemin = $translated; Resultat = ('erreur: ' + $_.Exception.Message) }) | Out-Null
            }
            if ($i % 50 -eq 0) { Write-Log ("... {0}/{1}" -f $i, $copied.Count) }
        }

        # ---- Elements des listes classiques ----
        $listResults = Join-Path $WorkDirectory 'SharePoint-Copy\resultats_elements_listes.csv'
        $listsFile = Join-Path $DiscoveryDirectory 'lists.json'
        if ((Test-Path -LiteralPath $listResults) -and (Test-Path -LiteralPath $listsFile)) {
            $discoveredLists = @((Get-Content -LiteralPath $listsFile -Raw -Encoding UTF8 | ConvertFrom-Json) | ForEach-Object { $_ })
            $listIdByKey = @{}
            foreach ($dl in $discoveredLists) { $listIdByKey[[string]$dl.relativeUrl] = [string]$dl.listId }

            $itemRows = @(Import-Csv -LiteralPath $listResults -Encoding UTF8 | Where-Object { $_.Status -eq 'copied' -and $_.TargetItemId })
            Write-Log ("Elements de listes copies a traiter : {0}" -f $itemRows.Count)
            $metaCache = @{}
            foreach ($r in $itemRows) {
                $key = [string]$r.ListKey
                if (-not $metaCache.ContainsKey($key)) {
                    $map = @{}
                    if ($listIdByKey.ContainsKey($key)) {
                        foreach ($row in (Read-Jsonl -Path (Join-Path $DiscoveryDirectory ("listitems\{0}.jsonl" -f $listIdByKey[$key])))) {
                            if ($row.PSObject.Properties['meta'] -and $row.meta) { $map[[int]$row.id] = $row.meta }
                        }
                    }
                    $metaCache[$key] = $map
                }
                $meta = $metaCache[$key][[int]$r.SourceItemId]
                if (-not $meta) { $stats.skipped++; continue }
                $label = "{0}/ID:{1}" -f $key, $r.TargetItemId
                if (-not $Apply) { $null = Set-Metadata -List $null -ItemId 0 -Meta $meta -Conn $null -Label $label; continue }

                try {
                    $targetKey = if ($LibraryNameMapping.ContainsKey($key)) { [string]$LibraryNameMapping[$key] } else { $key }
                    $list = $listByRel[$targetKey.ToLowerInvariant()]
                    if (-not $list) { throw "liste cible introuvable : $targetKey" }
                    $res = Set-Metadata -List $list -ItemId ([int]$r.TargetItemId) -Meta $meta -Conn $dstConn -Label $label
                    $results.Add([pscustomobject]@{ Type = 'Element'; Chemin = $label; Resultat = $res }) | Out-Null
                } catch {
                    $stats.failed++
                    Write-Log ("{0} : {1}" -f $label, $_.Exception.Message) 'ERROR'
                    $results.Add([pscustomobject]@{ Type = 'Element'; Chemin = $label; Resultat = ('erreur: ' + $_.Exception.Message) }) | Out-Null
                }
            }
        }

        Write-Log ("Metadonnees : {0} {1}, partielles (utilisateur introuvable) {2}, ecarts apres verification {3}, ignorees {4}, erreurs {5}" -f $(if ($Apply) { 'appliquees' } else { 'a appliquer (simulation)' }), $stats.ok, $stats.datesOnly, $stats.mismatch, $stats.skipped, $stats.failed) 'SUCCESS'
        if ($stats.failed -gt 0 -or $stats.mismatch -gt 0) { $script:PhaseFailed = $true }
    }
    catch {
        Write-Log ("Erreur fatale: {0}" -f $_.Exception.Message) 'ERROR'
        $script:PhaseFailed = $true
    }
    finally {
        if ($results.Count -gt 0) {
            $csv = Join-Path $OutputDirectory 'metadonnees_resultats.csv'
            $results | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
            Write-Log ("Rapport: {0}" -f $csv)
        }
    }
}

# ============================================================================
# ORCHESTRATION
# ============================================================================
$phaseFunctions = [ordered]@{
    Decouverte   = 'Invoke-PhaseDecouverte'
    Copie        = 'Invoke-PhaseCopie'
    Applications = 'Invoke-PhaseApplications'
    Navigation   = 'Invoke-PhaseNavigation'
    Droits       = 'Invoke-PhaseDroits'
    Metadonnees  = 'Invoke-PhaseMetadonnees'   # en dernier : les autres phases ne doivent plus modifier les elements apres elle
}

function Write-Banner([string]$Text, [string]$Color = 'Magenta') {
    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor $Color
    Write-Host $Text -ForegroundColor $Color
    Write-Host ('=' * 78) -ForegroundColor $Color
}

$overallStart = Get-Date
$phaseSummary = New-Object System.Collections.Generic.List[object]
$script:PhaseFailed = $false

Write-Banner ("MIGRATION SHAREPOINT - phases : {0} - mode PnP : {1}" -f ($Phases -join ', '), $(if ($Apply) { 'APPLICATION' } else { 'SIMULATION' }))
Write-Host "Source : $SourceSiteUrl"
Write-Host "Cible  : $TargetSiteUrl"

foreach ($name in $phaseFunctions.Keys) {
    if ($Phases -notcontains $name) { continue }

    Write-Banner ("PHASE : {0}" -f $name)
    $script:PhaseFailed = $false
    $started = Get-Date
    try {
        & $phaseFunctions[$name]
    } catch {
        Write-Host ("[ERREUR] Phase {0} interrompue : {1}" -f $name, $_.Exception.Message) -ForegroundColor Red
        $script:PhaseFailed = $true
    }

    $status = if ($script:PhaseFailed) { 'ECHEC' } else { 'OK' }
    $phaseSummary.Add([pscustomobject]@{ Phase = $name; Statut = $status; Duree = ((Get-Date) - $started).ToString('hh\:mm\:ss') }) | Out-Null

    if ($script:PhaseFailed -and -not $ContinueOnPhaseError) {
        Write-Host "Arret : la phase $name a echoue (ContinueOnPhaseError = false)." -ForegroundColor Yellow
        break
    }
}

Write-Banner ("RESUME - duree totale {0}" -f ((Get-Date) - $overallStart).ToString('hh\:mm\:ss')) 'Cyan'
$phaseSummary | Format-Table -AutoSize | Out-String | Write-Host

if (@($phaseSummary | Where-Object { $_.Statut -eq 'ECHEC' }).Count -gt 0) { exit 1 }
