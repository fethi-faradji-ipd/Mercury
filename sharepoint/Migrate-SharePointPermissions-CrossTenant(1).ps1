#requires -Version 7.4
#requires -Modules PnP.PowerShell

<#
.SYNOPSIS
    Migre les autorisations SharePoint entre deux sites de tenants differents.

.DESCRIPTION
    - Exporte les niveaux d'autorisation personnalises.
    - Exporte les groupes SharePoint et leurs membres.
    - Exporte les autorisations uniques du site, des listes/bibliotheques,
      des dossiers, des fichiers et des elements de liste.
    - Recree les niveaux et groupes dans le site cible.
    - Cree les utilisateurs externes manquants comme invites Entra B2B,
      sans envoyer de message d'invitation.
    - Reapplique les autorisations en faisant correspondre les objets par chemin.
    - Produit un journal CSV, un rapport d'invitations et un rapport des comptes
      ou groupes non resolus.
    - Utilise exclusivement une authentification applicative par certificat :
      une application et un certificat distincts pour chaque tenant.

    Tant que $Apply vaut $false dans la zone de configuration, aucune
    modification n'est effectuee dans le tenant cible.

.IMPORTANT
    Les listes et documents doivent deja exister dans le site cible. Les chemins
    relatifs des bibliotheques, dossiers et fichiers doivent etre identiques.
    Pour les listes classiques, le script suppose que les ID des elements ont ete
    preserves. Les groupes Entra/Microsoft 365 absents du tenant cible ne sont pas
    recrees automatiquement : renseignez $PrincipalMappingCsv pour les faire correspondre.
    Les liens de partage anonymes ou nominatifs ne sont pas recrees.

    Les deux certificats, avec leur cle privee, doivent etre installes dans le
    magasin de certificats Windows du compte qui execute le script. Aucune
    authentification interactive, aucun secret applicatif et aucun compte
    utilisateur ne sont utilises pendant l'execution.

    Autorisations applicatives a consentir par un administrateur :
    - application source : SharePoint Sites.FullControl.All, ou Sites.Selected
      avec FullControl sur le site source ;
    - application cible : SharePoint Sites.FullControl.All, ou Sites.Selected
      avec FullControl sur le site cible ;
    - application cible, Microsoft Graph : User.Read.All et User.Invite.All.

    La partie publique de chaque certificat doit etre ajoutee a l'application
    Entra correspondante. Le fichier PFX avec cle privee reste sur la machine.

.EXAMPLE
    # Apres avoir modifie la zone CONFIGURATION ci-dessous :
    .\Migrate-SharePointPermissions-CrossTenant.ps1
#>

# ============================================================================
# CONFIGURATION A MODIFIER
# ============================================================================

$SourceSiteUrl = 'https://TENANT-SOURCE.sharepoint.com/sites/NOM-DU-SITE'
$TargetSiteUrl = 'https://TENANT-CIBLE.sharepoint.com/sites/NOM-DU-SITE'

# Identifiant du tenant : GUID du tenant ou domaine tenant.onmicrosoft.com.
$SourceTenantId = '00000000-0000-0000-0000-000000000000'
$SourceClientId = '00000000-0000-0000-0000-000000000000'
$SourceCertificateThumbprint = 'EMPREINTE-CERTIFICAT-SOURCE'

$TargetTenantId = '11111111-1111-1111-1111-111111111111'
$TargetClientId = '11111111-1111-1111-1111-111111111111'
$TargetCertificateThumbprint = 'EMPREINTE-CERTIFICAT-CIBLE'

# Les certificats doivent etre places dans Cert:\CurrentUser\My ou
# Cert:\LocalMachine\My et comporter leur cle privee.

# $false = simulation sans modification. Mettre $true apres controle des rapports.
$Apply = $false

# $true = remplace les droits uniques des listes, dossiers, fichiers et elements.
# Les droits racine du site restent fusionnes pour eviter un verrouillage.
$ReplaceUniquePermissions = $false

# Mettre $true uniquement si les listes techniques/cachees doivent aussi etre traitees.
$IncludeHiddenLists = $false

# Mettre $true pour ne pas analyser les droits individuels des dossiers/fichiers/elements.
$SkipItemPermissions = $false

# Laisser vide tant qu'aucun fichier de correspondance n'est necessaire.
# Format CSV attendu : SourceKey,TargetLogin
$PrincipalMappingCsv = ''

# Les rapports seront crees dans un sous-dossier place a cote du script.
$OutputDirectory = Join-Path $PSScriptRoot 'SharePoint-Permissions-Migration'

# ============================================================================
# FIN DE LA CONFIGURATION - NE PAS MODIFIER LA SUITE
# ============================================================================

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

    $color = switch ($Status) {
        'Error'   { 'Red' }
        'Skipped' { 'Yellow' }
        'Planned' { 'Cyan' }
        default   { 'Green' }
    }
    Write-Host "[$Stage][$Status] $Action - $Object $Detail" -ForegroundColor $color
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
        Context       = $Context
        Reason        = $Reason
    })
    Add-MigrationLog -Stage 'Resolution' -Action 'Principal non resolu' -Status 'Skipped' -Object $Principal.Title -Detail $Reason
}

function Get-SafeFileName {
    param([Parameter(Mandatory)][string]$Value)
    return ($Value -replace '[^a-zA-Z0-9._-]', '_')
}

function Get-PrincipalRecord {
    param(
        [Parameter(Mandatory)]$Principal,
        [Parameter(Mandatory)]$Connection
    )

    $principalType = [string]$Principal.PrincipalType
    $loginName = [string]$Principal.LoginName
    $title = [string]$Principal.Title
    $email = ''

    if ($principalType -eq 'User') {
        try {
            Get-PnPProperty -ClientObject $Principal -Property Email -Connection $Connection | Out-Null
            $email = [string]$Principal.Email
        }
        catch {
            # Certains principaux systeme n'exposent pas l'adresse e-mail.
        }

        if ([string]::IsNullOrWhiteSpace($email) -and $title -match '^[^@\s]+@[^@\s]+$') {
            $email = $title
        }
        if ([string]::IsNullOrWhiteSpace($email)) {
            $loginCandidate = ($loginName -split '\|')[-1]
            if ($loginCandidate -notmatch '#EXT#' -and $loginCandidate -match '^[^@\s]+@[^@\s]+$') {
                $email = $loginCandidate
            }
        }
    }

    if ($principalType -eq 'User' -and -not [string]::IsNullOrWhiteSpace($email)) {
        $key = 'USER:' + $email.Trim().ToLowerInvariant()
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
        IsCustom        = ([string]$Role.RoleTypeKind -eq 'None' -and -not [bool]$Role.Hidden)
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
            Get-PnPProperty -ClientObject $role -Property Hidden -Connection $Connection | Out-Null
            if ([bool]$role.Hidden) {
                continue
            }
            $roles.Add([pscustomobject]@{
                Name         = [string]$role.Name
                RoleTypeKind = [string]$role.RoleTypeKind
                IsCustom     = ([string]$role.RoleTypeKind -eq 'None')
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

function Get-AssociatedGroupIds {
    param([Parameter(Mandatory)]$Connection)

    $ids = @{ Owners = $null; Members = $null; Visitors = $null }
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
                $ids[$association] = $group.Id
            }
        }
        catch {
            Add-MigrationLog -Stage 'Export' -Action 'Lecture groupe associe' -Status 'Skipped' -Object $association -Detail $_.Exception.Message
        }
    }
    return $ids
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
            $nextUri = [Uri]$nextLink
            $nextUrl = $nextUri.PathAndQuery
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

    Add-MigrationLog -Stage 'Export' -Action 'Lecture du site source' -Status 'Success' -Object $SourceSiteUrl
    $web = Get-PnPWeb -Includes ServerRelativeUrl, RoleAssignments -Connection $Connection

    $roleDefinitions = [System.Collections.Generic.List[object]]::new()
    foreach ($role in Get-PnPRoleDefinition -Connection $Connection) {
        $roleRecord = Get-RoleRecord -Role $role -Connection $Connection
        if ($roleRecord.IsCustom) {
            $roleDefinitions.Add($roleRecord)
        }
    }

    $associatedIds = Get-AssociatedGroupIds -Connection $Connection
    $groups = [System.Collections.Generic.List[object]]::new()
    foreach ($group in Get-PnPGroup -Connection $Connection) {
        Get-PnPProperty -ClientObject $group -Property Description, LoginName, AllowMembersEditMembership, OnlyAllowMembersViewMembership -Connection $Connection | Out-Null

        if ([string]$group.LoginName -like 'SharingLinks.*' -or [string]$group.Title -like 'SharingLinks.*') {
            Add-MigrationLog -Stage 'Export' -Action 'Lien de partage non exporte' -Status 'Skipped' -Object $group.Title -Detail 'Les liens de partage ne peuvent pas etre transposes entre tenants.'
            continue
        }

        $association = 'None'
        foreach ($key in $associatedIds.Keys) {
            if ($associatedIds[$key] -eq $group.Id) {
                $association = $key
                break
            }
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
            Association                    = $association
            AllowMembersEditMembership     = [bool]$group.AllowMembersEditMembership
            OnlyAllowMembersViewMembership = [bool]$group.OnlyAllowMembersViewMembership
            Members                        = @($members)
        })
    }

    $securableObjects = [System.Collections.Generic.List[object]]::new()
    $securableObjects.Add([pscustomobject]@{
        ObjectType      = 'Web'
        RelativePath   = ''
        ListRelativeUrl = $null
        ItemId         = $null
        FileRef        = $null
        IsFolder       = $false
        Assignments    = @(Get-Assignments -SecurableObject $web -Connection $Connection)
    })

    $lists = Get-PnPList -Includes RootFolder, Hidden, IsSystemList, HasUniqueRoleAssignments, BaseType, ItemCount -Connection $Connection
    $listIndex = 0
    foreach ($list in $lists) {
        $listIndex++
        Write-Progress -Activity 'Export des autorisations SharePoint' -Status $list.Title -PercentComplete (($listIndex / [Math]::Max(1, $lists.Count)) * 100)

        if (-not $IncludeHiddenLists -and ([bool]$list.Hidden -or [bool]$list.IsSystemList)) {
            continue
        }

        $listRelativeUrl = ([string]$list.RootFolder.ServerRelativeUrl).Substring(([string]$web.ServerRelativeUrl).Length).TrimStart('/')
        if ([bool]$list.HasUniqueRoleAssignments) {
            $securableObjects.Add([pscustomobject]@{
                ObjectType       = 'List'
                RelativePath    = $listRelativeUrl
                ListRelativeUrl = $listRelativeUrl
                ItemId          = $null
                FileRef         = $null
                IsFolder        = $false
                Assignments     = @(Get-Assignments -SecurableObject $list -Connection $Connection)
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
                    $fileRef.Substring(([string]$web.ServerRelativeUrl).Length).TrimStart('/')
                }
                else {
                    "$listRelativeUrl|ID:$($row.Id)"
                }

                $securableObjects.Add([pscustomobject]@{
                    ObjectType       = 'Item'
                    RelativePath    = $relativePath
                    ListRelativeUrl = $listRelativeUrl
                    ItemId          = [int]$row.Id
                    FileRef         = $fileRef
                    IsFolder        = ([int]$row.FileSystemObjectType -eq 1)
                    Assignments     = @(Get-Assignments -SecurableObject $item -Connection $Connection)
                })
            }
            catch {
                Add-MigrationLog -Stage 'Export' -Action 'Lecture droits element' -Status 'Error' -Object "$listRelativeUrl / ID $($row.Id)" -Detail $_.Exception.Message
            }
        }
    }
    Write-Progress -Activity 'Export des autorisations SharePoint' -Completed

    $export = [pscustomobject]@{
        SchemaVersion       = 1
        ExportedAtUtc       = (Get-Date).ToUniversalTime().ToString('o')
        SourceSiteUrl       = $SourceSiteUrl
        SourceWebRelativeUrl = [string]$web.ServerRelativeUrl
        CustomRoles         = @($roleDefinitions)
        SharePointGroups    = @($groups)
        SecurableObjects    = @($securableObjects)
    }

    $export | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $ExportFile -Encoding utf8
    Add-MigrationLog -Stage 'Export' -Action 'Export JSON termine' -Status 'Success' -Object $ExportFile -Detail "$($securableObjects.Count) objets securises"
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

function Get-GraphUserByEmail {
    param(
        [Parameter(Mandatory)][string]$Email,
        [Parameter(Mandatory)]$Connection
    )

    $escapedEmail = $Email.Replace("'", "''")
    $filter = "mail eq '$escapedEmail' or userPrincipalName eq '$escapedEmail'"
    $encodedFilter = [Uri]::EscapeDataString($filter)
    $response = Invoke-PnPGraphMethod -Method Get -Url "users?`$select=id,displayName,mail,userPrincipalName,userType,externalUserState&`$filter=$encodedFilter" -Connection $Connection
    $user = @($response.value) | Select-Object -First 1
    if ($null -ne $user) {
        return $user
    }

    # Certains invites ont l'adresse d'origine uniquement dans otherMails.
    $otherMailFilter = [Uri]::EscapeDataString("otherMails/any(address:address eq '$escapedEmail')")
    $response = Invoke-PnPGraphMethod -Method Get `
        -Url "users?`$select=id,displayName,mail,userPrincipalName,userType,externalUserState,otherMails&`$count=true&`$filter=$otherMailFilter" `
        -ConsistencyLevelEventual -Connection $Connection
    return @($response.value) | Select-Object -First 1
}

function Resolve-TargetUser {
    param(
        [Parameter(Mandatory)]$Principal,
        [Parameter(Mandatory)]$Connection,
        [Parameter(Mandatory)][string]$Context
    )

    if ($script:PrincipalMappings.ContainsKey([string]$Principal.Key)) {
        return $script:PrincipalMappings[[string]$Principal.Key]
    }
    if ($script:TargetPrincipalCache.ContainsKey([string]$Principal.Key)) {
        return $script:TargetPrincipalCache[[string]$Principal.Key]
    }

    if ([string]$Principal.PrincipalType -ne 'User' -or [string]::IsNullOrWhiteSpace([string]$Principal.Email)) {
        try {
            $matches = @(Get-PnPUser -Connection $Connection | Where-Object {
                $_.Title -eq [string]$Principal.Title -and $_.PrincipalType -ne 'SharePointGroup'
            })
            if ($matches.Count -eq 1) {
                $script:TargetPrincipalCache[[string]$Principal.Key] = [string]$matches[0].LoginName
                return [string]$matches[0].LoginName
            }
        }
        catch {
            # Le rapport detaille sera produit ci-dessous.
        }

        Add-UnresolvedPrincipal -Principal $Principal -Context $Context -Reason 'Principal sans adresse e-mail ou groupe Entra absent/ambigu. Ajouter une correspondance CSV.'
        return $null
    }

    $email = ([string]$Principal.Email).Trim().ToLowerInvariant()
    try {
        $targetUser = Get-GraphUserByEmail -Email $email -Connection $Connection
    }
    catch {
        Add-UnresolvedPrincipal -Principal $Principal -Context $Context -Reason "Recherche Graph impossible : $($_.Exception.Message)"
        return $null
    }

    if ($null -eq $targetUser) {
        if (-not $Apply) {
            Add-MigrationLog -Stage 'Invitation' -Action 'Creer invite sans e-mail' -Status 'Planned' -Object $email -Detail 'sendInvitationMessage=false'
            $script:TargetPrincipalCache[[string]$Principal.Key] = $email
            return $email
        }

        try {
            $invitation = Invoke-PnPGraphMethod -Method Post -Url 'invitations' -Content @{
                invitedUserEmailAddress = $email
                inviteRedirectUrl       = $TargetSiteUrl
                sendInvitationMessage   = $false
            } -Connection $Connection

            $targetUser = $invitation.invitedUser
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
            return $null
        }
    }

    $login = if (-not [string]::IsNullOrWhiteSpace([string]$targetUser.userPrincipalName)) {
        [string]$targetUser.userPrincipalName
    }
    else {
        $email
    }

    if ($Apply) {
        $resolvedInSharePoint = $false
        $lastError = $null
        for ($attempt = 1; $attempt -le 6 -and -not $resolvedInSharePoint; $attempt++) {
            try {
                New-PnPUser -LoginName $login -Connection $Connection | Out-Null
                $resolvedInSharePoint = $true
            }
            catch {
                $lastError = $_.Exception.Message
                if ($attempt -lt 6) {
                    Start-Sleep -Seconds 5
                }
            }
        }
        if (-not $resolvedInSharePoint) {
            Add-UnresolvedPrincipal -Principal $Principal -Context $Context -Reason "Utilisateur present dans Entra mais non resolu par SharePoint : $lastError"
            return $null
        }
    }

    $script:TargetPrincipalCache[[string]$Principal.Key] = $login
    return $login
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
        else {
            if ($Apply) {
                Set-PnPRoleDefinition -Identity $existing -Description $customRole.Description -ClearAll -Connection $Connection | Out-Null
                if (@($customRole.PermissionKinds).Count -gt 0) {
                    Set-PnPRoleDefinition -Identity $existing -Select @($customRole.PermissionKinds) -Connection $Connection | Out-Null
                }
                Add-MigrationLog -Stage 'Roles' -Action 'Niveau personnalise synchronise' -Status 'Success' -Object $customRole.Name
            }
            else {
                Add-MigrationLog -Stage 'Roles' -Action 'Synchroniser niveau personnalise' -Status 'Planned' -Object $customRole.Name
            }
        }
        $script:TargetRoleCache["NAME:$($customRole.Name)"] = $customRole.Name
    }
}

function Resolve-TargetRoleName {
    param([Parameter(Mandatory)]$Role)

    if ([bool]$Role.IsCustom -and $script:TargetRoleCache.ContainsKey("NAME:$($Role.Name)")) {
        return $script:TargetRoleCache["NAME:$($Role.Name)"]
    }
    if ([string]$Role.RoleTypeKind -ne 'None' -and $script:TargetRoleCache.ContainsKey("TYPE:$($Role.RoleTypeKind)")) {
        return $script:TargetRoleCache["TYPE:$($Role.RoleTypeKind)"]
    }
    if ($script:TargetRoleCache.ContainsKey("NAME:$($Role.Name)")) {
        return $script:TargetRoleCache["NAME:$($Role.Name)"]
    }
    return $null
}

function Initialize-TargetGroups {
    param(
        [Parameter(Mandatory)]$Export,
        [Parameter(Mandatory)]$Connection
    )

    $targetGroups = @(Get-PnPGroup -Connection $Connection)
    $targetAssociated = @{}
    $associatedSwitches = @{
        Owners   = 'AssociatedOwnerGroup'
        Members  = 'AssociatedMemberGroup'
        Visitors = 'AssociatedVisitorGroup'
    }
    foreach ($association in $associatedSwitches.Keys) {
        try {
            $parameters = @{ Connection = $Connection }
            $parameters[$associatedSwitches[$association]] = $true
            $targetAssociated[$association] = Get-PnPGroup @parameters
        }
        catch {
            $targetAssociated[$association] = $null
        }
    }

    foreach ($sourceGroup in @($Export.SharePointGroups)) {
        $sourceKey = 'SPGROUP:' + [string]$sourceGroup.Title
        $targetGroup = $null

        if ([string]$sourceGroup.Association -ne 'None') {
            $targetGroup = $targetAssociated[[string]$sourceGroup.Association]
        }
        if ($null -eq $targetGroup) {
            $targetGroup = $targetGroups | Where-Object Title -eq $sourceGroup.Title | Select-Object -First 1
        }

        if ($null -eq $targetGroup) {
            if ($Apply) {
                $targetGroup = New-PnPGroup -Title $sourceGroup.Title -Description $sourceGroup.Description -Connection $Connection
                Set-PnPGroup -Identity $targetGroup `
                    -AllowMembersEditMembership ([bool]$sourceGroup.AllowMembersEditMembership) `
                    -OnlyAllowMembersViewMembership ([bool]$sourceGroup.OnlyAllowMembersViewMembership) `
                    -Connection $Connection | Out-Null
                Add-MigrationLog -Stage 'Groups' -Action 'Groupe SharePoint cree' -Status 'Success' -Object $sourceGroup.Title
            }
            else {
                Add-MigrationLog -Stage 'Groups' -Action 'Creer groupe SharePoint' -Status 'Planned' -Object $sourceGroup.Title
                $script:TargetGroupMap[$sourceKey] = [string]$sourceGroup.Title
                continue
            }
        }

        $script:TargetGroupMap[$sourceKey] = [string]$targetGroup.Title
    }

    foreach ($sourceGroup in @($Export.SharePointGroups)) {
        $sourceKey = 'SPGROUP:' + [string]$sourceGroup.Title
        if (-not $script:TargetGroupMap.ContainsKey($sourceKey)) {
            continue
        }
        $targetGroupTitle = $script:TargetGroupMap[$sourceKey]

        foreach ($member in @($sourceGroup.Members)) {
            $context = "Groupe SharePoint : $($sourceGroup.Title)"
            $login = Resolve-TargetUser -Principal $member -Connection $Connection -Context $context
            if ([string]::IsNullOrWhiteSpace($login)) {
                continue
            }
            if ($Apply) {
                try {
                    Add-PnPGroupMember -Group $targetGroupTitle -LoginName $login -Connection $Connection
                    Add-MigrationLog -Stage 'Groups' -Action 'Membre ajoute' -Status 'Success' -Object $targetGroupTitle -Detail $login
                }
                catch {
                    if ($_.Exception.Message -match 'already exists|deja|déjà') {
                        Add-MigrationLog -Stage 'Groups' -Action 'Membre deja present' -Status 'Success' -Object $targetGroupTitle -Detail $login
                    }
                    else {
                        Add-MigrationLog -Stage 'Groups' -Action 'Ajout membre' -Status 'Error' -Object $targetGroupTitle -Detail $_.Exception.Message
                    }
                }
            }
            else {
                Add-MigrationLog -Stage 'Groups' -Action 'Ajouter membre' -Status 'Planned' -Object $targetGroupTitle -Detail $login
            }
        }
    }
}

function Resolve-AssignmentPrincipal {
    param(
        [Parameter(Mandatory)]$Principal,
        [Parameter(Mandatory)]$Connection,
        [Parameter(Mandatory)][string]$Context
    )

    if ($script:PrincipalMappings.ContainsKey([string]$Principal.Key)) {
        return [pscustomobject]@{ Kind = 'User'; Value = $script:PrincipalMappings[[string]$Principal.Key] }
    }
    if ([string]$Principal.PrincipalType -eq 'SharePointGroup') {
        if ($script:TargetGroupMap.ContainsKey([string]$Principal.Key)) {
            return [pscustomobject]@{ Kind = 'Group'; Value = $script:TargetGroupMap[[string]$Principal.Key] }
        }
        Add-UnresolvedPrincipal -Principal $Principal -Context $Context -Reason 'Groupe SharePoint cible introuvable.'
        return $null
    }

    $login = Resolve-TargetUser -Principal $Principal -Connection $Connection -Context $Context
    if ([string]::IsNullOrWhiteSpace($login)) {
        return $null
    }
    return [pscustomobject]@{ Kind = 'User'; Value = $login }
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

    if (-not [string]::IsNullOrWhiteSpace([string]$SourceObject.FileRef)) {
        $targetFileRef = [string]$SourceObject.FileRef
        if ($targetFileRef.StartsWith($SourceWebRelativeUrl, [StringComparison]::OrdinalIgnoreCase)) {
            $targetFileRef = $TargetWebRelativeUrl + $targetFileRef.Substring($SourceWebRelativeUrl.Length)
        }
        $escaped = $targetFileRef.Replace("'", "''")
        $listId = $TargetList.Id.ToString()
        $url = "/_api/web/lists(guid'$listId')/items?`$select=Id&`$filter=FileRef eq '$escaped'&`$top=2"
        $response = Invoke-PnPSPRestMethod -Method Get -Url $url -Connection $Connection
        $matches = @($response.value)
        if ($matches.Count -eq 1) {
            return [int]$matches[0].Id
        }
        return $null
    }

    # Pour une liste classique, on suppose que l'outil de migration a conserve l'ID.
    try {
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

    if (-not $Apply -or -not $ReplaceUniquePermissions) {
        return
    }

    try {
        Get-PnPProperty -ClientObject $SecurableObject -Property HasUniqueRoleAssignments -Connection $Connection | Out-Null
        if ([bool]$SecurableObject.HasUniqueRoleAssignments) {
            $SecurableObject.ResetRoleInheritance()
            Invoke-PnPQuery -Connection $Connection
        }
        $SecurableObject.BreakRoleInheritance($false, $false)
        Invoke-PnPQuery -Connection $Connection
        Add-MigrationLog -Stage 'Permissions' -Action 'Droits uniques reinitialises' -Status 'Success' -Object $Context
    }
    catch {
        throw "Impossible de reinitialiser les droits uniques de '$Context' : $($_.Exception.Message)"
    }
}

function Ensure-TargetListUniquePermissions {
    param(
        [Parameter(Mandatory)]$List,
        [Parameter(Mandatory)]$Connection,
        [Parameter(Mandatory)][string]$Context
    )

    if ($ReplaceUniquePermissions) {
        return
    }

    Get-PnPProperty -ClientObject $List -Property HasUniqueRoleAssignments -Connection $Connection | Out-Null
    if ([bool]$List.HasUniqueRoleAssignments) {
        return
    }

    if ($Apply) {
        $List.BreakRoleInheritance($true, $false)
        Invoke-PnPQuery -Connection $Connection
        Add-MigrationLog -Stage 'Permissions' -Action 'Heritage rompu avec copie des droits parents' -Status 'Success' -Object $Context
    }
    else {
        Add-MigrationLog -Stage 'Permissions' -Action 'Rompre heritage avec copie des droits parents' -Status 'Planned' -Object $Context
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
            $common.SystemUpdate = $true
            Set-PnPListItemPermission @common | Out-Null
        }
    }
    Add-MigrationLog -Stage 'Permissions' -Action 'Autorisation ajoutee' -Status 'Success' -Object $Context -Detail "$($ResolvedPrincipal.Value) -> $RoleName"
}

function Import-SharePointSecurity {
    param(
        [Parameter(Mandatory)]$Export,
        [Parameter(Mandatory)]$Connection
    )

    Import-PrincipalMappings
    $targetWeb = Get-PnPWeb -Includes ServerRelativeUrl, RoleAssignments -Connection $Connection
    $targetWebRelativeUrl = [string]$targetWeb.ServerRelativeUrl
    $sourceWebRelativeUrl = [string]$Export.SourceWebRelativeUrl

    Initialize-TargetRoles -Export $Export -Connection $Connection
    Initialize-TargetGroups -Export $Export -Connection $Connection
    $targetLists = Get-TargetListMap -Connection $Connection -TargetWebRelativeUrl $targetWebRelativeUrl

    $objectIndex = 0
    foreach ($sourceObject in @($Export.SecurableObjects)) {
        $objectIndex++
        Write-Progress -Activity 'Import des autorisations SharePoint' -Status $sourceObject.RelativePath -PercentComplete (($objectIndex / [Math]::Max(1, @($Export.SecurableObjects).Count)) * 100)

        $context = switch ([string]$sourceObject.ObjectType) {
            'Web'  { 'Site racine' }
            'List' { "Liste : $($sourceObject.ListRelativeUrl)" }
            default { "Element : $($sourceObject.RelativePath)" }
        }

        $targetList = $null
        $targetItemId = 0
        $targetSecurableObject = $null

        if ([string]$sourceObject.ObjectType -in @('List', 'Item')) {
            $listKey = ([string]$sourceObject.ListRelativeUrl).ToLowerInvariant()
            if (-not $targetLists.ContainsKey($listKey)) {
                Add-MigrationLog -Stage 'Permissions' -Action 'Objet cible introuvable' -Status 'Skipped' -Object $context -Detail 'Bibliotheque ou liste absente.'
                continue
            }
            $targetList = $targetLists[$listKey]
        }

        if ([string]$sourceObject.ObjectType -eq 'List') {
            $targetSecurableObject = $targetList
        }
        elseif ([string]$sourceObject.ObjectType -eq 'Item') {
            $targetItemId = Find-TargetItemId -SourceObject $sourceObject -TargetList $targetList `
                -SourceWebRelativeUrl $sourceWebRelativeUrl -TargetWebRelativeUrl $targetWebRelativeUrl -Connection $Connection
            if ($null -eq $targetItemId) {
                Add-MigrationLog -Stage 'Permissions' -Action 'Element cible introuvable' -Status 'Skipped' -Object $context -Detail 'Chemin ou ID non retrouve.'
                continue
            }
            $targetSecurableObject = Get-PnPListItem -List $targetList -Id $targetItemId -Connection $Connection
        }

        $resolvedAssignments = [System.Collections.Generic.List[object]]::new()
        foreach ($assignment in @($sourceObject.Assignments)) {
            $resolvedPrincipal = Resolve-AssignmentPrincipal -Principal $assignment.Principal -Connection $Connection -Context $context
            if ($null -eq $resolvedPrincipal) {
                continue
            }

            $roleNames = [System.Collections.Generic.List[string]]::new()
            foreach ($role in @($assignment.Roles)) {
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
            Add-MigrationLog -Stage 'Permissions' -Action 'Aucune autorisation applicable' -Status 'Skipped' -Object $context
            continue
        }

        if ([string]$sourceObject.ObjectType -eq 'List') {
            Ensure-TargetListUniquePermissions -List $targetList -Connection $Connection -Context $context
        }
        if ([string]$sourceObject.ObjectType -ne 'Web') {
            Clear-TargetUniquePermissions -SecurableObject $targetSecurableObject -Connection $Connection -Context $context
        }

        foreach ($assignment in $resolvedAssignments) {
            foreach ($roleName in $assignment.Roles) {
                try {
                    Add-TargetPermission -ObjectType ([string]$sourceObject.ObjectType) -TargetList $targetList `
                        -TargetItemId $targetItemId -ResolvedPrincipal $assignment.Principal -RoleName $roleName `
                        -Connection $Connection -Context $context
                }
                catch {
                    Add-MigrationLog -Stage 'Permissions' -Action 'Ajout autorisation' -Status 'Error' -Object $context -Detail $_.Exception.Message
                }
            }
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

    Write-Host "`nRapports :" -ForegroundColor Cyan
    Write-Host "- $logPath"
    Write-Host "- $unresolvedPath"
    Write-Host "- $invitationsPath"
}

function Assert-Configuration {
    foreach ($siteSetting in @(
            @{ Name = 'SourceSiteUrl'; Value = $SourceSiteUrl },
            @{ Name = 'TargetSiteUrl'; Value = $TargetSiteUrl }
        )) {
        $parsedUri = $null
        if ($siteSetting.Value -match 'TENANT-|NOM-DU-SITE' -or
            -not [Uri]::TryCreate([string]$siteSetting.Value, [UriKind]::Absolute, [ref]$parsedUri) -or
            $parsedUri.Scheme -ne 'https' -or
            $parsedUri.Host -notlike '*.sharepoint.com') {
            throw "Configuration invalide pour $($siteSetting.Name) : remplacez la valeur d'exemple par l'URL SharePoint complete."
        }
    }

    foreach ($clientSetting in @(
            @{ Name = 'SourceClientId'; Value = $SourceClientId; Placeholder = '00000000-0000-0000-0000-000000000000' },
            @{ Name = 'TargetClientId'; Value = $TargetClientId; Placeholder = '11111111-1111-1111-1111-111111111111' }
        )) {
        $parsedClientId = [guid]::Empty
        if ($clientSetting.Value -eq $clientSetting.Placeholder -or
            -not [guid]::TryParse([string]$clientSetting.Value, [ref]$parsedClientId)) {
            throw "Configuration invalide pour $($clientSetting.Name) : indiquez l'ID d'application Entra correspondant."
        }
    }

    if ($SourceClientId -eq $TargetClientId) {
        throw 'SourceClientId et TargetClientId doivent correspondre a deux applications distinctes.'
    }

    foreach ($tenantSetting in @(
            @{ Name = 'SourceTenantId'; Value = $SourceTenantId; Placeholder = '00000000-0000-0000-0000-000000000000' },
            @{ Name = 'TargetTenantId'; Value = $TargetTenantId; Placeholder = '11111111-1111-1111-1111-111111111111' }
        )) {
        if ([string]::IsNullOrWhiteSpace([string]$tenantSetting.Value) -or
            $tenantSetting.Value -eq $tenantSetting.Placeholder -or
            ([string]$tenantSetting.Value -notmatch '^[0-9a-fA-F-]{36}$' -and
             [string]$tenantSetting.Value -notmatch '^[A-Za-z0-9.-]+\.onmicrosoft\.com$')) {
            throw "Configuration invalide pour $($tenantSetting.Name) : indiquez le GUID du tenant ou son domaine tenant.onmicrosoft.com."
        }
    }

    if ($SourceTenantId -eq $TargetTenantId) {
        throw 'SourceTenantId et TargetTenantId doivent correspondre a deux tenants distincts.'
    }

    if (-not $IsWindows) {
        throw "L'authentification par empreinte exige Windows et le magasin de certificats Windows."
    }

    foreach ($certificateSetting in @(
            @{ Name = 'SourceCertificateThumbprint'; Value = $SourceCertificateThumbprint; Placeholder = 'EMPREINTE-CERTIFICAT-SOURCE' },
            @{ Name = 'TargetCertificateThumbprint'; Value = $TargetCertificateThumbprint; Placeholder = 'EMPREINTE-CERTIFICAT-CIBLE' }
        )) {
        $thumbprint = ([string]$certificateSetting.Value -replace '\s', '').ToUpperInvariant()
        if ($thumbprint -eq $certificateSetting.Placeholder -or $thumbprint -notmatch '^[0-9A-F]{40}$') {
            throw "Configuration invalide pour $($certificateSetting.Name) : indiquez l'empreinte du certificat."
        }

        $certificate = Get-ChildItem -Path Cert:\CurrentUser\My, Cert:\LocalMachine\My -ErrorAction SilentlyContinue |
            Where-Object Thumbprint -eq $thumbprint |
            Select-Object -First 1

        if ($null -eq $certificate) {
            throw "Certificat introuvable pour $($certificateSetting.Name) dans Cert:\CurrentUser\My ou Cert:\LocalMachine\My."
        }
        if (-not $certificate.HasPrivateKey) {
            throw "Le certificat $($certificateSetting.Name) ne comporte pas de cle privee. Importez le fichier PFX, pas seulement le CER."
        }
        if ($certificate.NotAfter -le (Get-Date)) {
            throw "Le certificat $($certificateSetting.Name) a expire le $($certificate.NotAfter)."
        }
    }

    $normalizedSourceThumbprint = ($SourceCertificateThumbprint -replace '\s', '').ToUpperInvariant()
    $normalizedTargetThumbprint = ($TargetCertificateThumbprint -replace '\s', '').ToUpperInvariant()
    if ($normalizedSourceThumbprint -eq $normalizedTargetThumbprint) {
        throw 'Les certificats source et cible doivent etre distincts.'
    }
}

if ($PSVersionTable.PSVersion -lt [version]'7.4') {
    throw 'PnP.PowerShell necessite PowerShell 7.4 ou une version ulterieure. Lancez le script avec pwsh, pas PowerShell ISE.'
}

Assert-Configuration
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$exportFile = Join-Path $OutputDirectory 'SharePointPermissions.json'
$sourceConnection = $null
$targetConnection = $null

try {
    Write-Host 'Connexion applicative au site source avec le certificat source...' -ForegroundColor Cyan
    $sourceConnection = Connect-PnPOnline `
        -Url $SourceSiteUrl `
        -Tenant $SourceTenantId `
        -ClientId $SourceClientId `
        -Thumbprint (($SourceCertificateThumbprint -replace '\s', '').ToUpperInvariant()) `
        -ReturnConnection
    $export = Export-SharePointSecurity -Connection $sourceConnection -ExportFile $exportFile

    Write-Host "`nConnexion applicative au site cible avec le certificat cible..." -ForegroundColor Cyan
    $targetConnection = Connect-PnPOnline `
        -Url $TargetSiteUrl `
        -Tenant $TargetTenantId `
        -ClientId $TargetClientId `
        -Thumbprint (($TargetCertificateThumbprint -replace '\s', '').ToUpperInvariant()) `
        -ReturnConnection

    if (-not $Apply) {
        Write-Host "`nMODE SIMULATION : aucune modification ne sera effectuee." -ForegroundColor Yellow
    }
    elseif ($ReplaceUniquePermissions) {
        Write-Host "`nMODE APPLICATION : les droits uniques des listes et elements seront remplaces." -ForegroundColor Yellow
    }
    else {
        Write-Host "`nMODE APPLICATION : les autorisations seront ajoutees/fusionnees." -ForegroundColor Yellow
    }

    Import-SharePointSecurity -Export $export -Connection $targetConnection
}
catch {
    Add-MigrationLog -Stage 'General' -Action 'Execution interrompue' -Status 'Error' -Object '' -Detail $_.Exception.Message
    throw
}
finally {
    Save-Reports -Directory $OutputDirectory
    # Disconnect-PnPOnline ne permet pas de cibler une connexion precise.
    # La liberation des variables suffit pour les connexions retournees par -ReturnConnection.
    $sourceConnection = $null
    $targetConnection = $null
}
