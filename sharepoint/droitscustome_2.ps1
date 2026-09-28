#requires -Version 5.1

<#
.COMPATIBILITE
    Utilisable depuis PowerShell ISE (Windows PowerShell 5.1) :
    - si PowerShell 7 (pwsh.exe) est installe, le script se relance automatiquement
      dedans (recommande, avec PnP.PowerShell 2.x/3.x installe pour pwsh) ;
    - sinon il s'execute en 5.1 avec PnP.PowerShell 1.12.0 (derniere version
      compatible, obsolete) : Install-Module PnP.PowerShell -RequiredVersion 1.12.0 -Scope CurrentUser
    Dans ISE, enregistrez le fichier et lancez-le avec F5 (pas F8).
#>

<#
.SYNOPSIS
    Migre les autorisations SharePoint NON NATIVES entre deux sites de tenants differents.

.DESCRIPTION
    Perimetre : uniquement ce qui a ete cree/ajoute manuellement.
    - Niveaux d'autorisation personnalises (les niveaux natifs sont ignores,
      y compris ceux dont RoleTypeKind vaut None comme "Affichage seul").
    - Groupes SharePoint personnalises, leurs membres, proprietaire et parametres
      (les groupes Proprietaires/Membres/Visiteurs du site, les groupes systeme
      et les liens de partage ne sont pas recrees).
    - Autorisations accordees directement a des utilisateurs, a des groupes
      Entra/Microsoft 365 et aux groupes personnalises, sur le site, les
      listes/bibliotheques, dossiers, fichiers et elements.
    - Les comptes systeme (SHAREPOINT\system, principaux d'application, etc.)
      sont toujours ignores.

    Tant que $Apply vaut $false, aucune modification n'est effectuee dans le tenant cible.

.IMPORTANT
    Le contenu doit deja exister dans le site cible avec les memes chemins relatifs.
    Pour les listes classiques (hors bibliotheques), le script suppose que les ID
    des elements ont ete preserves par l'outil de migration.
    Les liens de partage ne sont pas recrees tels quels : leurs destinataires
    (externes par defaut, voir $SharingLinkRecipients) recoivent un droit direct
    equivalent sur l'objet cible, sans aucun e-mail envoye.

    Autorisations applicatives a consentir par un administrateur :
    - application source : SharePoint Sites.FullControl.All (ou Sites.Selected FullControl) ;
    - application cible  : SharePoint Sites.FullControl.All (ou Sites.Selected FullControl) ;
    - application cible, Microsoft Graph : User.Read.All, GroupMember.Read.All,
      et User.Invite.All si $InviteMissingUsers = $true.

.EXAMPLE
    .\Migrate-SharePointPermissions-CrossTenant.ps1
#>

# ============================================================================
# CONFIGURATION A MODIFIER
# ============================================================================

$SourceSiteUrl = 'https://infoprodigital365.sharepoint.com/sites/fethygate'
$TargetSiteUrl = 'https://infoprodigital365.sharepoint.com/sites/fethiMercury'

# Identifiant du tenant : GUID du tenant ou domaine tenant.onmicrosoft.com.
$SourceTenantId = '2eea08b8-1972-447b-ad43-d044d042500a'
$SourceClientId = '821109a4-6e0e-48e8-b477-8f9b70aa32a4'
$SourceCertificatePath = 'C:\certs\MSGraphExchangeOnlineAuth20260717_2.pfx'
$SourceCertificatePassword = 'MotDePasseFort123!'

$TargetTenantId = '2eea08b8-1972-447b-ad43-d044d042500a'
$TargetClientId = '821109a4-6e0e-48e8-b477-8f9b70aa32a4'
$TargetCertificatePath = 'C:\certs\MSGraphExchangeOnlineAuth20260717_2.pfx'
$TargetCertificatePassword = 'MotDePasseFort123!'
# ATTENTION : les mots de passe sont stockes en clair dans ce fichier.
# Restreignez les droits NTFS du script au compte qui doit l'executer.

# $false = simulation sans modification. Mettre $true apres controle des rapports.
$Apply = $true

# $true = remplace les droits uniques des listes, dossiers, fichiers et elements.
# Les droits racine du site restent fusionnes pour eviter un verrouillage.
$ReplaceUniquePermissions = $false

# Mettre $true uniquement si les listes techniques/cachees doivent aussi etre traitees.
$IncludeHiddenLists = $false

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

# Les rapports seront crees dans un sous-dossier place a cote du script.
$OutputDirectory = Join-Path $(if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }) 'SharePoint-Permissions-Migration'

# Niveau de detail des logs a l'ecran et dans Suivi_<date>.log :
# 'Normal' = etapes, actions, avertissements et erreurs ;
# 'Detail' = en plus : chaque liste, element, principal et recherche Graph.
$LogLevel = 'Detail'

# ============================================================================
# FIN DE LA CONFIGURATION - NE PAS MODIFIER LA SUITE
# ============================================================================

# Un seul tenant : les parametres cible non renseignes reprennent ceux de la source.
if ([string]::IsNullOrWhiteSpace($TargetTenantId)) { $TargetTenantId = $SourceTenantId }
if ([string]::IsNullOrWhiteSpace($TargetClientId)) { $TargetClientId = $SourceClientId }
if ([string]::IsNullOrWhiteSpace($TargetCertificatePath)) { $TargetCertificatePath = $SourceCertificatePath }
if ([string]::IsNullOrWhiteSpace($TargetCertificatePassword)) { $TargetCertificatePassword = $SourceCertificatePassword }

# --- Demarrage : PowerShell ISE / Windows PowerShell 5.1 ou PowerShell 7 ---
$pnpModule = Get-Module -ListAvailable -Name PnP.PowerShell | Sort-Object Version -Descending | Select-Object -First 1

if ($PSVersionTable.PSVersion.Major -lt 7) {
    $pwshCommand = Get-Command -Name pwsh.exe -ErrorAction SilentlyContinue
    $useLegacyPnP = ($null -ne $pnpModule -and $pnpModule.Version.Major -lt 2)

    if ($null -ne $pwshCommand -and -not $useLegacyPnP) {
        $scriptPath = $PSCommandPath
        if ([string]::IsNullOrWhiteSpace($scriptPath)) {
            throw 'Enregistrez le script puis lancez-le avec F5 (l''execution d''une selection avec F8 n''est pas supportee).'
        }
        Write-Host "Windows PowerShell $($PSVersionTable.PSVersion) detecte : relance du script dans PowerShell 7 ($($pwshCommand.Source))..." -ForegroundColor Cyan
        & $pwshCommand.Source -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $scriptPath
        $childExitCode = $LASTEXITCODE
        if ($childExitCode -ne 0) {
            throw "Le script s'est termine en erreur dans PowerShell 7 (code $childExitCode). Consultez les messages ci-dessus et les rapports."
        }
        Write-Host 'Execution terminee dans PowerShell 7.' -ForegroundColor Green
        return
    }

    if (-not $useLegacyPnP) {
        throw @'
Ce script ne peut pas s'executer ici : PnP.PowerShell 2.x/3.x exige PowerShell 7.
Solution recommandee : installer PowerShell 7, puis PnP.PowerShell pour PowerShell 7 :
    winget install --id Microsoft.PowerShell --source winget
    pwsh -Command "Install-Module PnP.PowerShell -Scope CurrentUser"
Le script se relancera ensuite automatiquement depuis ISE.
Solution de repli (module obsolete), dans ISE :
    Install-Module PnP.PowerShell -RequiredVersion 1.12.0 -Scope CurrentUser
'@
    }

    Write-Warning "Execution sous Windows PowerShell $($PSVersionTable.PSVersion) avec PnP.PowerShell $($pnpModule.Version) (obsolete, non supporte). PowerShell 7 est recommande."
}
elseif ($null -eq $pnpModule) {
    throw 'Module PnP.PowerShell absent pour PowerShell 7 : Install-Module PnP.PowerShell -Scope CurrentUser'
}

Import-Module -Name PnP.PowerShell -RequiredVersion $pnpModule.Version -DisableNameChecking

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
    $script:ExitCode = 1
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

    # Lancement par double-clic / "Executer avec PowerShell" : garder la fenetre ouverte.
    $commandLine = [Environment]::GetCommandLineArgs()
    if ($Host.Name -eq 'ConsoleHost' -and ($commandLine -contains '-File' -or $commandLine -contains '-f') -and $commandLine -notcontains '-NonInteractive') {
        Read-Host "`nAppuyez sur Entree pour fermer la fenetre" | Out-Null
    }
}

if ((Get-Variable -Name ExitCode -Scope Script -ErrorAction SilentlyContinue) -and $script:ExitCode -ne 0) {
    exit $script:ExitCode
}
