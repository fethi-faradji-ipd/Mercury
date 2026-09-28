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

$DossierResultats = $PSScriptRoot
$NomRapportCSV = "Rapport-PowerBI-Consolide-90J.csv"
$NomCandidatsCSV = "Candidats-Retrait-Licence-PowerBI.csv"
$NomAuditDetailCSV = "Audit-PowerBI-Detail-90J.csv"
$NomConnexionsDetailCSV = "Connexions-PowerBI-Detail-30J.csv"
$NomFichierLog = "Rapport-PowerBI-90J.log"

$CheminRapportCSV = Join-Path $DossierResultats $NomRapportCSV
$CheminCandidatsCSV = Join-Path $DossierResultats $NomCandidatsCSV
$CheminAuditDetailCSV = Join-Path $DossierResultats $NomAuditDetailCSV
$CheminConnexionsDetailCSV = Join-Path $DossierResultats $NomConnexionsDetailCSV
$CheminLog = Join-Path $DossierResultats $NomFichierLog

$GraphBase = "https://graph.microsoft.com/v1.0"
$PowerBIServiceAppId = "00000009-0000-0000-c000-000000000000"
$PowerBIDesktopNames = @("Power BI Desktop", "Microsoft Power BI Desktop")

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

function Invoke-PowerBIRestApi {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [ValidateSet("GET", "POST")][string]$Method = "GET",
        [object]$Body = $null,
        [int]$MaxTentatives = 6
    )

    for ($Tentative = 1; $Tentative -le $MaxTentatives; $Tentative++) {
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
    $CorpsJeton = @{
        client_id = $ClientId
        client_secret = $ClientSecret
        scope = "https://graph.microsoft.com/.default"
        grant_type = "client_credentials"
    }
    $Jeton = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $CorpsJeton -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
    $script:EntetesGraph = @{ Authorization = "Bearer $($Jeton.access_token)" }
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

    $AuditPowerBI = @()
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

        Write-Log -Message "Recherche Purview terminee. Telechargement de tous les enregistrements." -Couleur Cyan
        $AuditPowerBI = Get-GraphCollection -Uri "$GraphBase/security/auditLog/queries/$IdRecherche/records?`$top=1000"
        $AuditComplet = $true
        Write-Log -Message "$($AuditPowerBI.Count) activite(s) Power BI recuperee(s) dans Purview." -Couleur Green
    }
    catch {
        $ErreurAudit = Get-ErrorMessage $_
        Write-Log -Message "Audit Purview incomplet : $ErreurAudit" -Niveau "ERREUR" -Couleur Red
        Write-Log -Message "Aucun retrait de licence ne sera conseille dans cet etat." -Niveau "SECURITE" -Couleur Red
    }

    $ConnexionsEntra = @()
    try {
        Write-Log -Message "Lecture des connexions Entra Power BI Online et Desktop (indice complementaire)." -Couleur Cyan
        $DateFiltreEntra = $DateDebutEntraUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
        # Les filtres Graph complexes avec OR et status/errorCode peuvent retourner HTTP 400.
        # L'endpoint signIns n'accepte pas non plus le parametre OData $select.
        # On interroge donc chaque application separement, puis on garde les succes localement.
        $FiltreOnline = [uri]::EscapeDataString("appId eq '$PowerBIServiceAppId' and createdDateTime ge $DateFiltreEntra")
        $UriOnline = "$GraphBase/auditLogs/signIns?`$filter=$FiltreOnline&`$top=1000"
        Write-Log -Message "Requete Entra 1/3 : Power BI Online."
        $OnlineBrut = Get-GraphCollection -Uri $UriOnline
        $Online = @($OnlineBrut | Where-Object { $null -ne $_.status -and [int64]$_.status.errorCode -eq 0 })

        $FiltreDesktop1 = [uri]::EscapeDataString("appDisplayName eq 'Power BI Desktop' and createdDateTime ge $DateFiltreEntra")
        $UriDesktop1 = "$GraphBase/auditLogs/signIns?`$filter=$FiltreDesktop1&`$top=1000"
        Write-Log -Message "Requete Entra 2/3 : Power BI Desktop."
        $Desktop1Brut = Get-GraphCollection -Uri $UriDesktop1

        $FiltreDesktop2 = [uri]::EscapeDataString("appDisplayName eq 'Microsoft Power BI Desktop' and createdDateTime ge $DateFiltreEntra")
        $UriDesktop2 = "$GraphBase/auditLogs/signIns?`$filter=$FiltreDesktop2&`$top=1000"
        Write-Log -Message "Requete Entra 3/3 : Microsoft Power BI Desktop."
        $Desktop2Brut = Get-GraphCollection -Uri $UriDesktop2

        $DesktopBrut = @($Desktop1Brut) + @($Desktop2Brut)
        $Desktop = @($DesktopBrut | Where-Object {
            $null -ne $_.status -and [int64]$_.status.errorCode -eq 0
        })

        $IdsConnexion = @{}
        $ListeConnexions = New-Object System.Collections.ArrayList
        foreach ($Connexion in @($Online) + @($Desktop)) {
            if ($null -eq $Connexion -or $IdsConnexion.ContainsKey([string]$Connexion.id)) { continue }
            $IdsConnexion[[string]$Connexion.id] = $true
            [void]$ListeConnexions.Add($Connexion)
        }
        $ConnexionsEntra = @($ListeConnexions)
        $ConnexionsEntraCompletes = $true
        Write-Log -Message "$($ConnexionsEntra.Count) connexion(s) Entra unique(s) recuperee(s)." -Couleur Green
    }
    catch {
        $ErreurConnexions = Get-ErrorMessage $_
        Write-Log -Message "Connexions Entra non disponibles : $ErreurConnexions" -Niveau "AVERTISSEMENT" -Couleur Yellow
    }

    $AuditParUpn = @{}
    foreach ($Evenement in $AuditPowerBI) {
        $Cle = $null
        if (-not [string]::IsNullOrWhiteSpace([string]$Evenement.userPrincipalName)) {
            $ClePossible = ([string]$Evenement.userPrincipalName).ToLowerInvariant()
            if ($UtilisateurParUpn.ContainsKey($ClePossible)) { $Cle = $ClePossible }
        }
        if ($null -eq $Cle -and $CleUtilisateurParId.ContainsKey([string]$Evenement.userId)) {
            $Cle = $CleUtilisateurParId[[string]$Evenement.userId]
        }
        if ($null -eq $Cle) {
            $IdentiteAudit = Get-AuditDataProperty $Evenement.auditData @("UserId", "UserKey", "UserPrincipalName")
            if (-not [string]::IsNullOrWhiteSpace($IdentiteAudit)) {
                $IdentiteAudit = $IdentiteAudit.ToLowerInvariant()
                if ($UtilisateurParUpn.ContainsKey($IdentiteAudit)) { $Cle = $IdentiteAudit }
                elseif ($CleUtilisateurParId.ContainsKey($IdentiteAudit)) { $Cle = $CleUtilisateurParId[$IdentiteAudit] }
            }
        }
        if ($null -eq $Cle) { continue }
        if (-not $UtilisateurParUpn.ContainsKey($Cle)) { continue }
        if (-not $AuditParUpn.ContainsKey($Cle)) { $AuditParUpn[$Cle] = New-Object System.Collections.ArrayList }
        [void]$AuditParUpn[$Cle].Add($Evenement)
    }

    $ConnexionParUpn = @{}
    foreach ($Connexion in $ConnexionsEntra) {
        $Cle = $null
        if (-not [string]::IsNullOrWhiteSpace([string]$Connexion.userPrincipalName)) {
            $ClePossible = ([string]$Connexion.userPrincipalName).ToLowerInvariant()
            if ($UtilisateurParUpn.ContainsKey($ClePossible)) { $Cle = $ClePossible }
        }
        if ($null -eq $Cle -and $CleUtilisateurParId.ContainsKey([string]$Connexion.userId)) {
            $Cle = $CleUtilisateurParId[[string]$Connexion.userId]
        }
        if ($null -eq $Cle) { continue }
        if (-not $UtilisateurParUpn.ContainsKey($Cle)) { continue }
        if (-not $ConnexionParUpn.ContainsKey($Cle)) { $ConnexionParUpn[$Cle] = New-Object System.Collections.ArrayList }
        [void]$ConnexionParUpn[$Cle].Add($Connexion)
    }

    $DetailAudit = foreach ($Cle in $AuditParUpn.Keys) {
        $Utilisateur = $UtilisateurParUpn[$Cle]
        foreach ($Evenement in $AuditParUpn[$Cle]) {
            [PSCustomObject]@{
                DateHeureLocale = Convert-ToLocalText $Evenement.createdDateTime
                DateHeureUTC = ([DateTimeOffset]$Evenement.createdDateTime).ToUniversalTime().ToString("o")
                Nom = $Utilisateur.DisplayName
                Utilisateur = $Utilisateur.UserPrincipalName
                LicencePowerBI = $Utilisateur.LicencePowerBI
                Operation = $Evenement.operation
                Service = $Evenement.service
                AdresseIP = $Evenement.clientIp
                TypeEnregistrement = $Evenement.auditLogRecordType
                EspaceDeTravail = Get-AuditDataProperty $Evenement.auditData @("WorkSpaceName", "WorkspaceName")
                Rapport = Get-AuditDataProperty $Evenement.auditData @("ReportName", "ItemName")
                JeuDeDonnees = Get-AuditDataProperty $Evenement.auditData @("DatasetName", "DataflowName")
                Objet = $Evenement.objectId
                IdEvenement = $Evenement.id
            }
        }
    }

    $DetailConnexions = foreach ($Cle in $ConnexionParUpn.Keys) {
        $Utilisateur = $UtilisateurParUpn[$Cle]
        foreach ($Connexion in $ConnexionParUpn[$Cle]) {
            $Canal = if ($PowerBIDesktopNames -contains [string]$Connexion.appDisplayName) { "Power BI Desktop" } else { "Power BI Online" }
            [PSCustomObject]@{
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
    }

    $Rapport = for ($Index = 0; $Index -lt $UtilisateursPayants.Count; $Index++) {
        $Utilisateur = $UtilisateursPayants[$Index]
        $Cle = $Utilisateur.UserPrincipalName.ToLowerInvariant()
        $Pourcentage = [math]::Round((($Index + 1) / $UtilisateursPayants.Count) * 100)
        Write-Progress -Activity "Consolidation de l'utilisation Power BI" -Status "$($Index + 1)/$($UtilisateursPayants.Count) - $($Utilisateur.UserPrincipalName)" -PercentComplete $Pourcentage

        $Activites = if ($AuditParUpn.ContainsKey($Cle)) { @($AuditParUpn[$Cle]) } else { @() }
        $Connexions = if ($ConnexionParUpn.ContainsKey($Cle)) { @($ConnexionParUpn[$Cle]) } else { @() }
        $DerniereActivite = $Activites | Sort-Object { [DateTimeOffset]$_.createdDateTime } -Descending | Select-Object -First 1
        $DerniereConnexion = $Connexions | Sort-Object { [DateTimeOffset]$_.createdDateTime } -Descending | Select-Object -First 1
        $DerniereOnline = $Connexions | Where-Object { $_.appId -eq $PowerBIServiceAppId } | Sort-Object { [DateTimeOffset]$_.createdDateTime } -Descending | Select-Object -First 1
        $DerniereDesktop = $Connexions | Where-Object { $PowerBIDesktopNames -contains [string]$_.appDisplayName } | Sort-Object { [DateTimeOffset]$_.createdDateTime } -Descending | Select-Object -First 1

        $Decision = "INDETERMINE - NE PAS RETIRER"
        $Motif = "Controle incomplet."
        $JoursDepuisActivite = $null
        if ($DerniereActivite) {
            $JoursDepuisActivite = [math]::Floor(($DateFinUtc - ([DateTimeOffset]$DerniereActivite.createdDateTime).UtcDateTime).TotalDays)
        }

        if (-not $AuditComplet) {
            $Motif = "Audit Purview incomplet ou en erreur : $ErreurAudit"
        }
        elseif ($DerniereActivite) {
            $Decision = "CONSERVER - UTILISEE"
            $Motif = "Activite Power BI reelle detectee dans Microsoft Purview Audit."
        }
        elseif (-not $ConnexionsEntraCompletes) {
            $Decision = "INDETERMINE - NE PAS RETIRER"
            $Motif = "Controle des connexions Entra incomplet : $ErreurConnexions"
        }
        elseif ($DerniereConnexion) {
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
            NombreActivitesAudit = $Activites.Count
            DerniereActivitePowerBI = if ($DerniereActivite) { Convert-ToLocalText $DerniereActivite.createdDateTime } else { $null }
            JoursDepuisDerniereActivite = $JoursDepuisActivite
            DerniereOperation = if ($DerniereActivite) { $DerniereActivite.operation } else { $null }
            ConnexionsEntraCompletes = $ConnexionsEntraCompletes
            DebutControleConnexions = Convert-ToLocalText $DateDebutEntraUtc
            NombreConnexionsEntra = $Connexions.Count
            DerniereConnexionEntra = if ($DerniereConnexion) { Convert-ToLocalText $DerniereConnexion.createdDateTime } else { $null }
            DerniereConnexionOnline = if ($DerniereOnline) { Convert-ToLocalText $DerniereOnline.createdDateTime } else { $null }
            DerniereConnexionDesktop = if ($DerniereDesktop) { Convert-ToLocalText $DerniereDesktop.createdDateTime } else { $null }
            ErreurAudit = $ErreurAudit
            ErreurConnexionsEntra = $ErreurConnexions
        }
    }
    Write-Progress -Activity "Consolidation de l'utilisation Power BI" -Completed

    $Rapport = @($Rapport | Sort-Object Decision, Utilisateur)
    $Candidats = @($Rapport | Where-Object { $_.Decision -eq "CANDIDAT AU RETRAIT - VALIDATION HUMAINE" })
    $DetailAudit = @($DetailAudit | Sort-Object DateHeureUTC -Descending)
    $DetailConnexions = @($DetailConnexions | Sort-Object DateHeureUTC -Descending)

    Write-Log -Message "Ecriture du rapport consolide : $CheminRapportCSV" -Couleur Cyan
    $Rapport | Export-Csv -LiteralPath $CheminRapportCSV -Delimiter ";" -NoTypeInformation -Encoding UTF8
    $Candidats | Export-Csv -LiteralPath $CheminCandidatsCSV -Delimiter ";" -NoTypeInformation -Encoding UTF8
    $DetailAudit | Export-Csv -LiteralPath $CheminAuditDetailCSV -Delimiter ";" -NoTypeInformation -Encoding UTF8
    $DetailConnexions | Export-Csv -LiteralPath $CheminConnexionsDetailCSV -Delimiter ";" -NoTypeInformation -Encoding UTF8

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
    Write-Host "Details Audit : $NomAuditDetailCSV" -ForegroundColor Green
    Write-Host "Details connexions : $NomConnexionsDetailCSV" -ForegroundColor Green
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
