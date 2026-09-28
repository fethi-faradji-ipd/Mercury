#requires -Version 5.1

<#
.SYNOPSIS
Rapport prudent des licences Power BI payantes et de leur utilisation sur 90 jours.

.DESCRIPTION
- Connexion non interactive avec une application Entra ID et un secret.
- Detection des licences contenant les plans BI_AZURE_P2 (Pro) ou BI_AZURE_P3 (PPU).
- Utilisation reelle recherchee dans Microsoft Purview Audit (powerBIAudit).
- Connexions Entra Power BI Service et Power BI Desktop ajoutees comme indices.
- Aucun utilisateur n'est propose au retrait si l'audit est incomplet ou en erreur.

Autorisations Microsoft Graph de type Application avec consentement administrateur :
  User.Read.All
  Organization.Read.All
  AuditLog.Read.All
  AuditLogsQuery.Read.All

Compatible Windows PowerShell 5.1 et PowerShell ISE. Aucun module requis.
#>

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# =============================================================================
# CONFIGURATION FIXE - A COMPLETER
# =============================================================================
$NomDuScript = "Rapport-PowerBI-90J-AppSecret-ISE.ps1"
$TenantId = "2eea08b8-1972-447b-ad43-d044d042500a"
$ClientId = "62a03bb6-b52a-4b2d-a3a7-82e542026446"
$ModeAuthentification = "Secret"


# Utilisé uniquement lorsque ModeAuthentification = 'Secret'.
# Saisissez la VALEUR du secret, et non l'identifiant du secret.
$ClientSecret = "6Xp8Q~PEE3L.1qRNR9EsSdV9.Adxi5mMFy_8TbgC"

$NombreJoursInactivite = 90
$NombreJoursConnexionsEntra = 30
$DelaiInterrogationSecondes = 30
$DureeMaxAttenteMinutes = 720
$InclureComptesDesactives = $false
$TaillePageFlux = 250
# Les fichiers detailles sont ecrits page par page, sans etre gardes en memoire.
# Mettre $false pour accelerer encore le script et ne produire que les 2 rapports principaux.
$GenererFichiersDetail = $true

$DossierResultats = $PSScriptRoot
$NomRapportCSV = "Rapport-PowerBI-Consolide-90J2.csv"
$NomCandidatsCSV = "Candidats-Retrait-Licence-PowerBI2.csv"
$NomAuditDetailCSV = "Audit-PowerBI-Detail-90J2.csv"
$NomConnexionsDetailCSV = "Connexions-PowerBI-Detail-30J2.csv"
$NomFichierLog = "Rapport-PowerBI-90J2.log"

$CheminRapportCSV = Join-Path $DossierResultats $NomRapportCSV
$CheminCandidatsCSV = Join-Path $DossierResultats $NomCandidatsCSV
$CheminAuditDetailCSV = Join-Path $DossierResultats $NomAuditDetailCSV
$CheminConnexionsDetailCSV = Join-Path $DossierResultats $NomConnexionsDetailCSV
$CheminLog = Join-Path $DossierResultats $NomFichierLog

$GraphBase = "https://graph.microsoft.com/v1.0"
$GraphBeta = "https://graph.microsoft.com/beta"
$PowerBIServiceAppId = "00000009-0000-0000-c000-000000000000"
$PowerBIDesktopNames = @("Power BI Desktop", "Microsoft Power BI Desktop", "Microsoft Power BI")

function Write-Log {
    param(
        [string]$Message,
        [string]$Niveau = "INFO",
        [ConsoleColor]$Couleur = [ConsoleColor]::Gray
    )
    $Horodatage = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $Ligne = "[$Horodatage] [$Niveau] $Message"
    Write-Host $Ligne -ForegroundColor $Couleur
    Add-Content -LiteralPath $CheminLog -Value $Ligne -Encoding UTF8
}

function Get-ErrorMessage {
    param([System.Management.Automation.ErrorRecord]$Erreur)
    $Message = $Erreur.Exception.Message
    try {
        if ($Erreur.ErrorDetails -and $Erreur.ErrorDetails.Message) {
            $Detail = $Erreur.ErrorDetails.Message | ConvertFrom-Json
            if ($Detail.error.message) {
                $Message = $Detail.error.message
            }
        }
    }
    catch { }
    return $Message
}

function Update-GraphAccessToken {
    Write-Log -Message "Obtention ou renouvellement du jeton Microsoft Graph." -Couleur Cyan
    $CorpsJeton = @{
        client_id = $ClientId
        client_secret = $ClientSecret
        scope = "https://graph.microsoft.com/.default"
        grant_type = "client_credentials"
    }
    $NouveauJeton = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $CorpsJeton -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace([string]$NouveauJeton.access_token)) {
        throw "Microsoft Entra n'a pas retourne de jeton d'acces."
    }
    $DureeJetonSecondes = [int]$NouveauJeton.expires_in
    if ($DureeJetonSecondes -le 0) { $DureeJetonSecondes = 3600 }
    $script:EntetesGraph = @{ Authorization = "Bearer $($NouveauJeton.access_token)" }
    # Renouvellement preventif cinq minutes avant l'expiration annoncee.
    $script:ExpirationJetonGraphUtc = (Get-Date).ToUniversalTime().AddSeconds($DureeJetonSecondes - 300)
    Write-Log -Message "Jeton Graph valide jusqu'au $($script:ExpirationJetonGraphUtc.ToLocalTime().ToString('dd/MM/yyyy HH:mm:ss'))." -Couleur Green
}

function Invoke-PowerBIRestApi {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [ValidateSet("GET", "POST")][string]$Method = "GET",
        [object]$Body = $null,
        [int]$MaxTentatives = 6
    )

    for ($Tentative = 1; $Tentative -le $MaxTentatives; $Tentative++) {
        if ($null -eq $script:ExpirationJetonGraphUtc -or (Get-Date).ToUniversalTime() -ge $script:ExpirationJetonGraphUtc) {
            Update-GraphAccessToken
        }
        try {
            $Parametres = @{
                Uri = $Uri
                Method = $Method
                Headers = $script:EntetesGraph
                ErrorAction = "Stop"
            }
            if ($null -ne $Body) {
                $Parametres.ContentType = "application/json"
                $Parametres.Body = $Body | ConvertTo-Json -Depth 20 -Compress
            }
            return Invoke-RestMethod @Parametres
        }
        catch {
            $Code = $null
            try { $Code = [int]$_.Exception.Response.StatusCode } catch { }
            if ($Code -eq 401 -and $Tentative -lt $MaxTentatives) {
                Write-Log -Message "Jeton Graph refuse ou expire (HTTP 401). Renouvellement immediat." -Niveau "AUTH" -Couleur Yellow
                Update-GraphAccessToken
                continue
            }
            $Reessayer = ($Code -eq 429 -or $Code -eq 500 -or $Code -eq 502 -or $Code -eq 503 -or $Code -eq 504)
            if (-not $Reessayer -or $Tentative -eq $MaxTentatives) { throw }
            $Pause = [math]::Min(60, [math]::Pow(2, $Tentative))
            Write-Log -Message "Graph retourne HTTP $Code. Nouvelle tentative dans $Pause seconde(s)." -Niveau "ATTENTE" -Couleur Yellow
            Start-Sleep -Seconds $Pause
        }
    }
}

function Get-GraphCollection {
    param([Parameter(Mandatory = $true)][string]$Uri)
    $Elements = New-Object System.Collections.ArrayList
    $Page = 0
    while (-not [string]::IsNullOrWhiteSpace($Uri)) {
        $Page++
        $Reponse = Invoke-PowerBIRestApi -Uri $Uri
        foreach ($Element in @($Reponse.value)) { [void]$Elements.Add($Element) }
        Write-Log -Message "Page $Page recue, total : $($Elements.Count) element(s)."
        $Uri = $Reponse.'@odata.nextLink'
    }
    return @($Elements)
}

function Invoke-GraphCollectionStreaming {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][scriptblock]$TraiterPage,
        [string]$Libelle = "Elements"
    )
    $Page = 0
    $Total = 0L
    while (-not [string]::IsNullOrWhiteSpace($Uri)) {
        $Page++
        $Reponse = Invoke-PowerBIRestApi -Uri $Uri
        $ElementsPage = @($Reponse.value)
        $Total += $ElementsPage.Count
        & $TraiterPage -Elements $ElementsPage
        $MemoireMo = [math]::Round((Get-Process -Id $PID).WorkingSet64 / 1MB)
        Write-Log -Message "$Libelle - page $Page traitee, total lu : $Total, memoire PowerShell : $MemoireMo Mo."
        $Uri = $Reponse.'@odata.nextLink'
        $ElementsPage = $null
        $Reponse = $null
        if (($Page % 20) -eq 0) {
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
    return $Total
}

function Convert-ToLocalText {
    param($Date)
    if ($null -eq $Date -or [string]::IsNullOrWhiteSpace([string]$Date)) { return $null }
    return ([DateTimeOffset]$Date).ToLocalTime().ToString("dd/MM/yyyy HH:mm:ss zzz")
}

function Get-AuditDataProperty {
    param($AuditData, [string[]]$Noms)
    if ($null -eq $AuditData) { return $null }
    $Objet = $AuditData
    if ($AuditData -is [string]) {
        try { $Objet = $AuditData | ConvertFrom-Json } catch { return $null }
    }
    foreach ($Nom in $Noms) {
        $Propriete = $Objet.PSObject.Properties[$Nom]
        if ($null -ne $Propriete -and $null -ne $Propriete.Value) { return [string]$Propriete.Value }
    }
    return $null
}

function Get-LicenseAssignmentDate {
    param($PlansPayants)
    $Dates = @($PlansPayants | Where-Object { $null -ne $_.AssignedDateTime } | ForEach-Object { [DateTimeOffset]$_.AssignedDateTime })
    if ($Dates.Count -eq 0) { return $null }
    return ($Dates | Sort-Object -Descending | Select-Object -First 1)
}

function Resolve-AuditUserKey {
    param($Evenement)
    if (-not [string]::IsNullOrWhiteSpace([string]$Evenement.userPrincipalName)) {
        $ClePossible = ([string]$Evenement.userPrincipalName).ToLowerInvariant()
        if ($script:UtilisateurParUpn.ContainsKey($ClePossible)) { return $ClePossible }
    }
    if ($script:CleUtilisateurParId.ContainsKey([string]$Evenement.userId)) {
        return $script:CleUtilisateurParId[[string]$Evenement.userId]
    }
    $IdentiteAudit = Get-AuditDataProperty $Evenement.auditData @("UserId", "UserKey", "UserPrincipalName")
    if (-not [string]::IsNullOrWhiteSpace($IdentiteAudit)) {
        $IdentiteAudit = $IdentiteAudit.ToLowerInvariant()
        if ($script:UtilisateurParUpn.ContainsKey($IdentiteAudit)) { return $IdentiteAudit }
        if ($script:CleUtilisateurParId.ContainsKey($IdentiteAudit)) { return $script:CleUtilisateurParId[$IdentiteAudit] }
    }
    return $null
}

function Resolve-SignInUserKey {
    param($Connexion)
    if (-not [string]::IsNullOrWhiteSpace([string]$Connexion.userPrincipalName)) {
        $ClePossible = ([string]$Connexion.userPrincipalName).ToLowerInvariant()
        if ($script:UtilisateurParUpn.ContainsKey($ClePossible)) { return $ClePossible }
    }
    if ($script:CleUtilisateurParId.ContainsKey([string]$Connexion.userId)) {
        return $script:CleUtilisateurParId[[string]$Connexion.userId]
    }
    return $null
}

function Update-AuditSummary {
    param([string]$Cle, $Evenement)
    if (-not $script:AuditResumeParUpn.ContainsKey($Cle)) {
        $script:AuditResumeParUpn[$Cle] = [PSCustomObject]@{
            Nombre = 0L
            DerniereDateUtc = $null
            DerniereOperation = $null
        }
    }
    $Resume = $script:AuditResumeParUpn[$Cle]
    $Resume.Nombre++
    $DateEvenement = ([DateTimeOffset]$Evenement.createdDateTime).UtcDateTime
    if ($null -eq $Resume.DerniereDateUtc -or $DateEvenement -gt $Resume.DerniereDateUtc) {
        $Resume.DerniereDateUtc = $DateEvenement
        $Resume.DerniereOperation = [string]$Evenement.operation
    }
}

function Update-SignInSummary {
    param([string]$Cle, $Connexion)
    if (-not $script:ConnexionResumeParUpn.ContainsKey($Cle)) {
        $script:ConnexionResumeParUpn[$Cle] = [PSCustomObject]@{
            Nombre = 0L
            DerniereDateUtc = $null
            DerniereOnlineUtc = $null
            DerniereDesktopUtc = $null
        }
    }
    $Resume = $script:ConnexionResumeParUpn[$Cle]
    $Resume.Nombre++
    $DateConnexion = ([DateTimeOffset]$Connexion.createdDateTime).UtcDateTime
    if ($null -eq $Resume.DerniereDateUtc -or $DateConnexion -gt $Resume.DerniereDateUtc) {
        $Resume.DerniereDateUtc = $DateConnexion
    }
    $NomApplication = [string]$Connexion.appDisplayName
    $EstDesktop = ($NomApplication -match "Desktop" -or $NomApplication -eq "Microsoft Power BI")
    if ($EstDesktop) {
        if ($null -eq $Resume.DerniereDesktopUtc -or $DateConnexion -gt $Resume.DerniereDesktopUtc) {
            $Resume.DerniereDesktopUtc = $DateConnexion
        }
    }
    else {
        if ($null -eq $Resume.DerniereOnlineUtc -or $DateConnexion -gt $Resume.DerniereOnlineUtc) {
            $Resume.DerniereOnlineUtc = $DateConnexion
        }
    }
}

function Export-AuditDetailPage {
    param([object[]]$Lignes)
    if (-not $GenererFichiersDetail -or $Lignes.Count -eq 0) { return }
    if (-not $script:AuditCsvCommence) {
        $Lignes | Export-Csv -LiteralPath $CheminAuditDetailCSV -Delimiter ";" -NoTypeInformation -Encoding UTF8
        $script:AuditCsvCommence = $true
    }
    else {
        $Lignes | Export-Csv -LiteralPath $CheminAuditDetailCSV -Delimiter ";" -NoTypeInformation -Encoding UTF8 -Append
    }
}

function Export-SignInDetailPage {
    param([object[]]$Lignes)
    if (-not $GenererFichiersDetail -or $Lignes.Count -eq 0) { return }
    if (-not $script:SignInCsvCommence) {
        $Lignes | Export-Csv -LiteralPath $CheminConnexionsDetailCSV -Delimiter ";" -NoTypeInformation -Encoding UTF8
        $script:SignInCsvCommence = $true
    }
    else {
        $Lignes | Export-Csv -LiteralPath $CheminConnexionsDetailCSV -Delimiter ";" -NoTypeInformation -Encoding UTF8 -Append
    }
}

function New-SignInDetailRow {
    param([string]$Cle, $Connexion)
    $Utilisateur = $script:UtilisateurParUpn[$Cle]
    $NomApplication = [string]$Connexion.appDisplayName
    $Canal = if ($NomApplication -match "Desktop" -or $NomApplication -eq "Microsoft Power BI") { "Power BI Desktop" } else { "Power BI Online" }
    return [PSCustomObject]@{
        DateHeureLocale = Convert-ToLocalText $Connexion.createdDateTime
        DateHeureUTC = ([DateTimeOffset]$Connexion.createdDateTime).ToUniversalTime().ToString("o")
        Nom = $Utilisateur.DisplayName
        Utilisateur = $Utilisateur.UserPrincipalName
        LicencePowerBI = $Utilisateur.LicencePowerBI
        Canal = $Canal
        Application = $Connexion.appDisplayName
        Interactif = $Connexion.isInteractive
        TypesConnexion = [string]::Join(" | ", [string[]]@($Connexion.signInEventTypes))
        AdresseIP = $Connexion.ipAddress
        Ressource = $Connexion.resourceDisplayName
        Systeme = $Connexion.deviceDetail.operatingSystem
        Navigateur = $Connexion.deviceDetail.browser
        Ville = $Connexion.location.city
        Pays = $Connexion.location.countryOrRegion
        IdEvenement = $Connexion.id
        CorrelationId = $Connexion.correlationId
    }
}

Set-Content -LiteralPath $CheminLog -Value "" -Encoding UTF8
$Chronometre = [Diagnostics.Stopwatch]::StartNew()
$AuditComplet = $false
$ConnexionsEntraCompletes = $false
$ErreurAudit = $null
$ErreurConnexions = $null

try {
    Write-Log -Message "Demarrage de $NomDuScript" -Couleur Cyan

    if ($TenantId -like "REMPLACER-*" -or $ClientId -like "REMPLACER-*" -or $ClientSecret -like "REMPLACER-*") {
        throw "Completez TenantId, ClientId et ClientSecret dans le bloc CONFIGURATION en haut du script."
    }
    if ($NombreJoursInactivite -lt 1) { throw "NombreJoursInactivite doit etre superieur a zero." }

    $DateFinUtc = (Get-Date).ToUniversalTime()
    $DateDebutAuditUtc = $DateFinUtc.AddDays(-$NombreJoursInactivite)
    $DateDebutEntraUtc = $DateFinUtc.AddDays(-$NombreJoursConnexionsEntra)
    $DateLimiteRetraitUtc = $DateFinUtc.AddDays(-$NombreJoursInactivite)

    Write-Log -Message "Periode Audit Power BI : du $(Convert-ToLocalText $DateDebutAuditUtc) au $(Convert-ToLocalText $DateFinUtc)." -Couleur Cyan
    Write-Log -Message "Periode Connexions Entra : du $(Convert-ToLocalText $DateDebutEntraUtc) au $(Convert-ToLocalText $DateFinUtc)." -Couleur Cyan

    Write-Log -Message "Connexion a Microsoft Graph avec l'application Azure et son secret." -Couleur Cyan
    Update-GraphAccessToken
    Write-Log -Message "Connexion applicative reussie." -Couleur Green

    Write-Log -Message "Lecture des abonnements et plans de service du tenant." -Couleur Cyan
    $Skus = Get-GraphCollection -Uri "$GraphBase/subscribedSkus?`$select=skuId,skuPartNumber,servicePlans"
    $PlansPowerBIPayants = @{}
    $NomSkuParId = @{}
    foreach ($Sku in $Skus) {
        $NomSkuParId[[string]$Sku.skuId] = [string]$Sku.skuPartNumber
        foreach ($Plan in @($Sku.servicePlans)) {
            if ($Plan.servicePlanName -in @("BI_AZURE_P2", "BI_AZURE_P3")) {
                $PlansPowerBIPayants[[string]$Plan.servicePlanId] = [string]$Plan.servicePlanName
            }
        }
    }
    if ($PlansPowerBIPayants.Count -eq 0) { throw "Aucun plan BI_AZURE_P2 ou BI_AZURE_P3 n'a ete trouve dans les abonnements du tenant." }

    Write-Log -Message "Lecture de tous les utilisateurs et de leurs licences." -Couleur Cyan
    $UriUtilisateurs = "$GraphBase/users?`$select=id,displayName,userPrincipalName,mail,accountEnabled,userType,assignedPlans,assignedLicenses&`$top=999"
    $TousUtilisateurs = Get-GraphCollection -Uri $UriUtilisateurs

    $UtilisateursPayants = New-Object System.Collections.ArrayList
    foreach ($Utilisateur in $TousUtilisateurs) {
        $PlansPayants = @($Utilisateur.assignedPlans | Where-Object {
            $_.capabilityStatus -eq "Enabled" -and $PlansPowerBIPayants.ContainsKey([string]$_.servicePlanId)
        })
        if ($PlansPayants.Count -eq 0) { continue }
        if (-not $InclureComptesDesactives -and -not $Utilisateur.accountEnabled) { continue }

        $TypesPlan = @($PlansPayants | ForEach-Object {
            if ($PlansPowerBIPayants[[string]$_.servicePlanId] -eq "BI_AZURE_P3") { "Power BI Premium par utilisateur" } else { "Power BI Pro" }
        } | Sort-Object -Unique)
        $SkusUtilisateur = @($Utilisateur.assignedLicenses | ForEach-Object { $NomSkuParId[[string]$_.skuId] } | Where-Object { $_ } | Sort-Object -Unique)
        $DateAttribution = Get-LicenseAssignmentDate -PlansPayants $PlansPayants

        [void]$UtilisateursPayants.Add([PSCustomObject]@{
            Id = $Utilisateur.id
            DisplayName = $Utilisateur.displayName
            UserPrincipalName = $Utilisateur.userPrincipalName
            Mail = $Utilisateur.mail
            AccountEnabled = $Utilisateur.accountEnabled
            UserType = $Utilisateur.userType
            LicencePowerBI = [string]::Join(" + ", [string[]]$TypesPlan)
            SkusAttribues = [string]::Join(" | ", [string[]]$SkusUtilisateur)
            DateAttribution = $DateAttribution
        })
    }
    $UtilisateursPayants = @($UtilisateursPayants | Sort-Object UserPrincipalName)
    Write-Log -Message "$($UtilisateursPayants.Count) utilisateur(s) avec Power BI payant trouve(s)." -Couleur Green
    if ($UtilisateursPayants.Count -eq 0) { throw "Aucun utilisateur Power BI payant actif n'a ete trouve." }

    $UtilisateurParUpn = @{}
    $CleUtilisateurParId = @{}
    foreach ($Utilisateur in $UtilisateursPayants) {
        if (-not [string]::IsNullOrWhiteSpace($Utilisateur.UserPrincipalName)) {
            $UtilisateurParUpn[$Utilisateur.UserPrincipalName.ToLowerInvariant()] = $Utilisateur
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$Utilisateur.Id)) {
            $CleUtilisateurParId[[string]$Utilisateur.Id] = $Utilisateur.UserPrincipalName.ToLowerInvariant()
        }
    }

    # Resumes compacts : quelques valeurs par utilisateur, jamais les evenements complets.
    $script:AuditResumeParUpn = @{}
    $script:ConnexionResumeParUpn = @{}
    $script:AuditCsvCommence = $false
    $script:SignInCsvCommence = $false
    if ($GenererFichiersDetail) {
        Set-Content -LiteralPath $CheminAuditDetailCSV -Value "" -Encoding UTF8
        Set-Content -LiteralPath $CheminConnexionsDetailCSV -Value "" -Encoding UTF8
    }

    # Les objets utilisateurs complets ne sont plus necessaires.
    $TousUtilisateurs = $null
    $Skus = $null
    [GC]::Collect()

    try {
        Write-Log -Message "Creation de la recherche Microsoft Purview Audit Power BI sur $NombreJoursInactivite jours." -Couleur Cyan
        $NomRecherche = "PowerBI-Licences-$($DateFinUtc.ToString('yyyyMMdd-HHmmss'))"
        $CorpsRecherche = @{
            displayName = $NomRecherche
            filterStartDateTime = $DateDebutAuditUtc.ToString("o")
            filterEndDateTime = $DateFinUtc.ToString("o")
            recordTypeFilters = @("powerBIAudit")
        }
        $Recherche = Invoke-PowerBIRestApi -Uri "$GraphBase/security/auditLog/queries" -Method POST -Body $CorpsRecherche
        $IdRecherche = [string]$Recherche.id
        if ([string]::IsNullOrWhiteSpace($IdRecherche)) { throw "Microsoft Graph n'a pas retourne l'identifiant de la recherche Purview." }
        Write-Log -Message "Recherche Purview creee. Identifiant : $IdRecherche" -Couleur Green

        $DebutAttente = Get-Date
        $StatutRecherche = [string]$Recherche.status
        while ($StatutRecherche -notin @("succeeded", "failed", "cancelled")) {
            $Minutes = [math]::Round(((Get-Date) - $DebutAttente).TotalMinutes, 1)
            if ($Minutes -ge $DureeMaxAttenteMinutes) { throw "La recherche Purview n'est pas terminee apres $DureeMaxAttenteMinutes minutes." }
            Write-Log -Message "Recherche Purview en cours. Statut : $StatutRecherche. Attente : $Minutes minute(s)." -Niveau "ATTENTE" -Couleur Yellow
            Start-Sleep -Seconds $DelaiInterrogationSecondes
            $Recherche = Invoke-PowerBIRestApi -Uri "$GraphBase/security/auditLog/queries/$IdRecherche"
            $StatutRecherche = [string]$Recherche.status
        }
        if ($StatutRecherche -ne "succeeded") { throw "La recherche Purview s'est terminee avec le statut : $StatutRecherche." }

        Write-Log -Message "Recherche Purview terminee. Traitement des enregistrements page par page." -Couleur Cyan
        $TraiterPageAudit = {
            param([object[]]$Elements)
            $LignesPage = New-Object System.Collections.ArrayList
            foreach ($Evenement in $Elements) {
                $Cle = Resolve-AuditUserKey $Evenement
                if ([string]::IsNullOrWhiteSpace($Cle)) { continue }
                Update-AuditSummary -Cle $Cle -Evenement $Evenement
                if ($GenererFichiersDetail) {
                    $UtilisateurDetail = $script:UtilisateurParUpn[$Cle]
                    [void]$LignesPage.Add([PSCustomObject]@{
                        DateHeureLocale = Convert-ToLocalText $Evenement.createdDateTime
                        DateHeureUTC = ([DateTimeOffset]$Evenement.createdDateTime).ToUniversalTime().ToString("o")
                        Nom = $UtilisateurDetail.DisplayName
                        Utilisateur = $UtilisateurDetail.UserPrincipalName
                        LicencePowerBI = $UtilisateurDetail.LicencePowerBI
                        Operation = $Evenement.operation
                        Service = $Evenement.service
                        AdresseIP = $Evenement.clientIp
                        TypeEnregistrement = $Evenement.auditLogRecordType
                        EspaceDeTravail = Get-AuditDataProperty $Evenement.auditData @("WorkSpaceName", "WorkspaceName")
                        Rapport = Get-AuditDataProperty $Evenement.auditData @("ReportName", "ItemName")
                        JeuDeDonnees = Get-AuditDataProperty $Evenement.auditData @("DatasetName", "DataflowName")
                        Objet = $Evenement.objectId
                        IdEvenement = $Evenement.id
                    })
                }
            }
            Export-AuditDetailPage -Lignes @($LignesPage)
        }
        $NombreAuditLu = Invoke-GraphCollectionStreaming -Uri "$GraphBase/security/auditLog/queries/$IdRecherche/records?`$top=$TaillePageFlux" -TraiterPage $TraiterPageAudit -Libelle "Audit Purview"
        $AuditComplet = $true
        Write-Log -Message "$NombreAuditLu activite(s) Power BI lue(s) dans Purview sans stockage en memoire." -Couleur Green
    }
    catch {
        $ErreurAudit = Get-ErrorMessage $_
        Write-Log -Message "Audit Purview incomplet : $ErreurAudit" -Niveau "ERREUR" -Couleur Red
        Write-Log -Message "Aucun retrait de licence ne sera conseille dans cet etat." -Niveau "SECURITE" -Couleur Red
    }

    $FiltreTypesConnexion = "signInEventTypes/any(t: t eq 'interactiveUser' or t eq 'nonInteractiveUser')"
    $DateFiltreEntra = $DateDebutEntraUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
    try {
        Write-Log -Message "Lecture des connexions Entra Power BI Online et Desktop (indice complementaire)." -Couleur Cyan
        # Deux prefixes non chevauchants couvrent notamment Power BI Service,
        # Power BI Desktop et Microsoft Power BI sans creer de doublons en memoire.
        $FiltrePowerBI = [uri]::EscapeDataString("$FiltreTypesConnexion and startsWith(appDisplayName,'Power BI') and createdDateTime ge $DateFiltreEntra")
        $UriPowerBI = "$GraphBeta/auditLogs/signIns?`$filter=$FiltrePowerBI&`$top=$TaillePageFlux"
        $FiltreMicrosoftPowerBI = [uri]::EscapeDataString("$FiltreTypesConnexion and startsWith(appDisplayName,'Microsoft Power BI') and createdDateTime ge $DateFiltreEntra")
        $UriMicrosoftPowerBI = "$GraphBeta/auditLogs/signIns?`$filter=$FiltreMicrosoftPowerBI&`$top=$TaillePageFlux"

        $TraiterPageConnexion = {
            param([object[]]$Elements)
            $LignesPage = New-Object System.Collections.ArrayList
            foreach ($Connexion in $Elements) {
                if ($null -eq $Connexion.status -or [int64]$Connexion.status.errorCode -ne 0) { continue }
                $Cle = Resolve-SignInUserKey $Connexion
                if ([string]::IsNullOrWhiteSpace($Cle)) { continue }
                Update-SignInSummary -Cle $Cle -Connexion $Connexion
                if ($GenererFichiersDetail) { [void]$LignesPage.Add((New-SignInDetailRow -Cle $Cle -Connexion $Connexion)) }
            }
            Export-SignInDetailPage -Lignes @($LignesPage)
        }

        Write-Log -Message "Requete Entra 1/2 : applications commencant par Power BI."
        $NombreConnexion1 = Invoke-GraphCollectionStreaming -Uri $UriPowerBI -TraiterPage $TraiterPageConnexion -Libelle "Connexions Power BI"
        Write-Log -Message "Requete Entra 2/2 : applications commencant par Microsoft Power BI."
        $NombreConnexion2 = Invoke-GraphCollectionStreaming -Uri $UriMicrosoftPowerBI -TraiterPage $TraiterPageConnexion -Libelle "Connexions Microsoft Power BI"
        $ConnexionsEntraCompletes = $true
        Write-Log -Message "$($NombreConnexion1 + $NombreConnexion2) connexion(s) Entra lue(s) sans stockage en memoire." -Couleur Green
    }
    catch {
        $ErreurConnexions = Get-ErrorMessage $_
        Write-Log -Message "Connexions Entra non disponibles : $ErreurConnexions" -Niveau "AVERTISSEMENT" -Couleur Yellow
    }

    # Controle de securite individuel par ID Entra.
    # Il est lance uniquement pour les utilisateurs sans activite Purview et sans
    # connexion Power BI deja rapprochee. Cela reproduit le filtre du portail Entra
    # sur un utilisateur precis et evite les erreurs de rapprochement par UPN.
    $VerificationConnexionParUpn = @{}
    $SourceVerificationParUpn = @{}
    $ErreurConnexionParUpn = @{}
    $UtilisateursAControler = @($UtilisateursPayants | Where-Object {
        $CleTest = $_.UserPrincipalName.ToLowerInvariant()
        -not $script:AuditResumeParUpn.ContainsKey($CleTest) -and -not $script:ConnexionResumeParUpn.ContainsKey($CleTest)
    })

    Write-Log -Message "$($UtilisateursAControler.Count) utilisateur(s) sans activite rapprochee : controle individuel par ID Entra." -Couleur Cyan

    foreach ($UtilisateurControle in $UtilisateursPayants) {
        $CleControle = $UtilisateurControle.UserPrincipalName.ToLowerInvariant()

        if ($script:AuditResumeParUpn.ContainsKey($CleControle)) {
            $VerificationConnexionParUpn[$CleControle] = $true
            $SourceVerificationParUpn[$CleControle] = "Non requise - activite Purview trouvee"
            continue
        }
        if ($script:ConnexionResumeParUpn.ContainsKey($CleControle)) {
            $VerificationConnexionParUpn[$CleControle] = $true
            $SourceVerificationParUpn[$CleControle] = "Recherche globale Entra"
            continue
        }

        try {
            Write-Log -Message "Controle individuel Entra : $($UtilisateurControle.UserPrincipalName) - ID $($UtilisateurControle.Id)"
            $FiltreIndividuelTexte = "$FiltreTypesConnexion and userId eq '$($UtilisateurControle.Id)' and createdDateTime ge $DateFiltreEntra"
            $FiltreIndividuel = [uri]::EscapeDataString($FiltreIndividuelTexte)
            $UriIndividuelle = "$GraphBeta/auditLogs/signIns?`$filter=$FiltreIndividuel&`$top=$TaillePageFlux"
            $NombreTrouveAvant = if ($script:ConnexionResumeParUpn.ContainsKey($CleControle)) { $script:ConnexionResumeParUpn[$CleControle].Nombre } else { 0L }
            $TraiterPageIndividuelle = {
                param([object[]]$Elements)
                $LignesPage = New-Object System.Collections.ArrayList
                foreach ($Connexion in $Elements) {
                    $Succes = ($null -ne $Connexion.status -and [int64]$Connexion.status.errorCode -eq 0)
                    $ApplicationPowerBI = (
                        $Connexion.appId -eq $PowerBIServiceAppId -or
                        [string]$Connexion.appDisplayName -match "Power\s*BI" -or
                        [string]$Connexion.resourceDisplayName -match "Power\s*BI"
                    )
                    if (-not ($Succes -and $ApplicationPowerBI)) { continue }
                    Update-SignInSummary -Cle $CleControle -Connexion $Connexion
                    if ($GenererFichiersDetail) { [void]$LignesPage.Add((New-SignInDetailRow -Cle $CleControle -Connexion $Connexion)) }
                }
                Export-SignInDetailPage -Lignes @($LignesPage)
            }
            $null = Invoke-GraphCollectionStreaming -Uri $UriIndividuelle -TraiterPage $TraiterPageIndividuelle -Libelle "Controle $($UtilisateurControle.UserPrincipalName)"
            $NombreTrouveApres = if ($script:ConnexionResumeParUpn.ContainsKey($CleControle)) { $script:ConnexionResumeParUpn[$CleControle].Nombre } else { 0L }
            $NombreConnexionsTrouvees = $NombreTrouveApres - $NombreTrouveAvant

            $VerificationConnexionParUpn[$CleControle] = $true
            $SourceVerificationParUpn[$CleControle] = "Controle individuel par ID Entra"

            if ($NombreConnexionsTrouvees -gt 0) {
                Write-Log -Message "$NombreConnexionsTrouvees connexion(s) Power BI trouvee(s) pour $($UtilisateurControle.UserPrincipalName)." -Couleur Green
            }
            else {
                Write-Log -Message "Aucune connexion Power BI trouvee pour $($UtilisateurControle.UserPrincipalName) pendant le controle individuel."
            }
        }
        catch {
            $MessageControle = Get-ErrorMessage $_
            $VerificationConnexionParUpn[$CleControle] = $false
            $SourceVerificationParUpn[$CleControle] = "Echec du controle individuel"
            $ErreurConnexionParUpn[$CleControle] = $MessageControle
            Write-Log -Message "Controle individuel impossible pour $($UtilisateurControle.UserPrincipalName) : $MessageControle" -Niveau "AVERTISSEMENT" -Couleur Yellow
        }
    }

    $Rapport = for ($Index = 0; $Index -lt $UtilisateursPayants.Count; $Index++) {
        $Utilisateur = $UtilisateursPayants[$Index]
        $Cle = $Utilisateur.UserPrincipalName.ToLowerInvariant()
        $Pourcentage = [math]::Round((($Index + 1) / $UtilisateursPayants.Count) * 100)
        Write-Progress -Activity "Consolidation de l'utilisation Power BI" -Status "$($Index + 1)/$($UtilisateursPayants.Count) - $($Utilisateur.UserPrincipalName)" -PercentComplete $Pourcentage

        $ResumeAudit = if ($script:AuditResumeParUpn.ContainsKey($Cle)) { $script:AuditResumeParUpn[$Cle] } else { $null }
        $ResumeConnexion = if ($script:ConnexionResumeParUpn.ContainsKey($Cle)) { $script:ConnexionResumeParUpn[$Cle] } else { $null }
        $DerniereActiviteUtc = if ($null -ne $ResumeAudit) { $ResumeAudit.DerniereDateUtc } else { $null }
        $DerniereConnexionUtc = if ($null -ne $ResumeConnexion) { $ResumeConnexion.DerniereDateUtc } else { $null }
        $DerniereOnlineUtc = if ($null -ne $ResumeConnexion) { $ResumeConnexion.DerniereOnlineUtc } else { $null }
        $DerniereDesktopUtc = if ($null -ne $ResumeConnexion) { $ResumeConnexion.DerniereDesktopUtc } else { $null }
        $ControleConnexionUtilisateurComplet = ($VerificationConnexionParUpn.ContainsKey($Cle) -and $VerificationConnexionParUpn[$Cle])
        $ErreurConnexionUtilisateur = if ($ErreurConnexionParUpn.ContainsKey($Cle)) { $ErreurConnexionParUpn[$Cle] } else { $ErreurConnexions }

        $Decision = "INDETERMINE - NE PAS RETIRER"
        $Motif = "Controle incomplet."
        $JoursDepuisActivite = $null
        if ($null -ne $DerniereActiviteUtc) {
            $JoursDepuisActivite = [math]::Floor(($DateFinUtc - $DerniereActiviteUtc).TotalDays)
        }

        if (-not $AuditComplet) {
            $Motif = "Audit Purview incomplet ou en erreur : $ErreurAudit"
        }
        elseif ($null -ne $DerniereActiviteUtc) {
            $Decision = "CONSERVER - UTILISEE"
            $Motif = "Activite Power BI reelle detectee dans Microsoft Purview Audit."
        }
        elseif (-not $ControleConnexionUtilisateurComplet) {
            $Decision = "INDETERMINE - NE PAS RETIRER"
            $Motif = "Controle individuel des connexions Entra incomplet : $ErreurConnexionUtilisateur"
        }
        elseif ($null -ne $DerniereConnexionUtc) {
            $Decision = "CONSERVER - CONNEXION RECENTE"
            $Motif = "Connexion Power BI Online ou Desktop detectee dans Entra. Par prudence, la licence est conservee."
        }
        elseif ($null -ne $Utilisateur.DateAttribution -and $Utilisateur.DateAttribution.UtcDateTime -gt $DateLimiteRetraitUtc) {
            $Decision = "CONSERVER - LICENCE TROP RECENTE"
            $Motif = "Licence ou plan payant attribue depuis moins de $NombreJoursInactivite jours."
        }
        elseif (-not $Utilisateur.AccountEnabled) {
            $Decision = "A VERIFIER - COMPTE DESACTIVE"
            $Motif = "Compte desactive avec licence payante encore attribuee."
        }
        else {
            $Decision = "CANDIDAT AU RETRAIT - VALIDATION HUMAINE"
            $Motif = "Aucune activite Power BI trouvee sur toute la periode de $NombreJoursInactivite jours. Valider avec le manager avant retrait."
        }

        [PSCustomObject]@{
            Decision = $Decision
            Motif = $Motif
            Nom = $Utilisateur.DisplayName
            Utilisateur = $Utilisateur.UserPrincipalName
            Mail = $Utilisateur.Mail
            CompteActif = $Utilisateur.AccountEnabled
            TypeUtilisateur = $Utilisateur.UserType
            LicencePowerBI = $Utilisateur.LicencePowerBI
            SkusAttribues = $Utilisateur.SkusAttribues
            DateAttributionLicence = Convert-ToLocalText $Utilisateur.DateAttribution
            AuditPurviewComplet = $AuditComplet
            DebutControleAudit = Convert-ToLocalText $DateDebutAuditUtc
            FinControleAudit = Convert-ToLocalText $DateFinUtc
            NombreActivitesAudit = if ($null -ne $ResumeAudit) { $ResumeAudit.Nombre } else { 0 }
            DerniereActivitePowerBI = Convert-ToLocalText $DerniereActiviteUtc
            JoursDepuisDerniereActivite = $JoursDepuisActivite
            DerniereOperation = if ($null -ne $ResumeAudit) { $ResumeAudit.DerniereOperation } else { $null }
            ConnexionsEntraCompletes = $ControleConnexionUtilisateurComplet
            SourceControleConnexions = $SourceVerificationParUpn[$Cle]
            DebutControleConnexions = Convert-ToLocalText $DateDebutEntraUtc
            NombreConnexionsEntra = if ($null -ne $ResumeConnexion) { $ResumeConnexion.Nombre } else { 0 }
            DerniereConnexionEntra = Convert-ToLocalText $DerniereConnexionUtc
            DerniereConnexionOnline = Convert-ToLocalText $DerniereOnlineUtc
            DerniereConnexionDesktop = Convert-ToLocalText $DerniereDesktopUtc
            ErreurAudit = $ErreurAudit
            ErreurConnexionsEntra = $ErreurConnexionUtilisateur
        }
    }
    Write-Progress -Activity "Consolidation de l'utilisation Power BI" -Completed

    $Rapport = @($Rapport | Sort-Object Decision, Utilisateur)
    $Candidats = @($Rapport | Where-Object { $_.Decision -eq "CANDIDAT AU RETRAIT - VALIDATION HUMAINE" })
    Write-Log -Message "Ecriture du rapport consolide : $CheminRapportCSV" -Couleur Cyan
    $Rapport | Export-Csv -LiteralPath $CheminRapportCSV -Delimiter ";" -NoTypeInformation -Encoding UTF8
    $Candidats | Export-Csv -LiteralPath $CheminCandidatsCSV -Delimiter ";" -NoTypeInformation -Encoding UTF8

    Write-Log -Message "Rapport termine : $($Rapport.Count) licence(s), $($Candidats.Count) candidat(s) au retrait." -Couleur Green
    if (-not $AuditComplet) {
        Write-Log -Message "ATTENTION : audit incomplet, aucun utilisateur ne doit etre decommissionne avec ce resultat." -Niveau "SECURITE" -Couleur Red
    }
    else {
        Write-Log -Message "Les candidats exigent encore une validation humaine avant retrait de licence." -Niveau "SECURITE" -Couleur Yellow
    }
    Write-Host ""
    Write-Host "Fichiers generes dans : $DossierResultats" -ForegroundColor Cyan
    Write-Host "Rapport principal : $NomRapportCSV" -ForegroundColor Green
    Write-Host "Candidats : $NomCandidatsCSV" -ForegroundColor Green
    if ($GenererFichiersDetail) {
        Write-Host "Details Audit : $NomAuditDetailCSV" -ForegroundColor Green
        Write-Host "Details connexions : $NomConnexionsDetailCSV" -ForegroundColor Green
    }
    Write-Host "Journal : $NomFichierLog" -ForegroundColor Green
}
catch {
    $MessageFatal = Get-ErrorMessage $_
    Write-Log -Message "Arret du script : $MessageFatal" -Niveau "FATAL" -Couleur Red
    Write-Host "Aucune licence ne doit etre retiree a partir d'une execution en erreur." -ForegroundColor Red
    throw
}
finally {
    $Chronometre.Stop()
    Write-Log -Message "Duree totale : $($Chronometre.Elapsed.ToString())."
}
