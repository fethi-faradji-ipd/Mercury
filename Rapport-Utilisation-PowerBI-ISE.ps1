#requires -Version 5.1

# VERSION ISE : connexion applicative Azure/Entra (AppOnly), Windows PowerShell 5.1

<#
.SYNOPSIS
Exporte un resume par utilisateur et l'historique detaille des authentifications
Power BI Online et Power BI Desktop sur les 30 derniers jours.

.EXAMPLE
.\Rapport-Utilisation-PowerBI-ISE-corrige.ps1

.NOTES
Une authentification ne prouve pas une utilisation active de Power BI : le SSO
silencieux et le renouvellement de jetons peuvent aussi produire des événements.

L'application Entra doit disposer des autorisations Microsoft Graph de type
Application User.Read.All et AuditLog.Read.All, avec consentement administrateur.
#>

$ErrorActionPreference = "Stop"

# ============================================================================
# CONFIGURATION DE L'APPLICATION AZURE / MICROSOFT ENTRA
# Remplacez les valeurs ci-dessous par celles de votre inscription d'application.
# ModeAuthentification accepte : 'Secret' ou 'Certificat'.
# ============================================================================
$NomDuScript = "Rapport-Utilisation-PowerBI-ISE-corrige.ps1"
$NomDuScript = "Rapport-Utilisation-PowerBI-ISE.ps1"
$TenantId = "2eea08b8-1972-447b-ad43-d044d042500a"
$ClientId = "62a03bb6-b52a-4b2d-a3a7-82e542026446"
$ModeAuthentification = "Secret"


# Utilisé uniquement lorsque ModeAuthentification = 'Secret'.
# Saisissez la VALEUR du secret, et non l'identifiant du secret.
$ClientSecretEnClair = "6Xp8Q~PEE3L.1qRNR9EsSdV9.Adxi5mMFy_8TbgC"


# Utilisé uniquement lorsque ModeAuthentification = 'Certificat'.
$CertificateThumbprint = "REMPLACER-PAR-EMPREINTE-DU-CERTIFICAT"

# Nom fixe du rapport genere dans le meme dossier que le script.
$NomFichierCSV = "Comparatif-Utilisation-PowerBI-Desktop-Online.csv"
$CheminCSV = Join-Path $PSScriptRoot $NomFichierCSV

# Historique detaille comparable aux journaux affiches dans Entra.
$NomFichierHistoriqueCSV = "Historique-Connexions-PowerBI-30-jours.csv"
$CheminHistoriqueCSV = Join-Path $PSScriptRoot $NomFichierHistoriqueCSV

# Entra P1/P2 conserve normalement les journaux de connexion pendant 30 jours.
$NombreJoursAnalyse = 30

# Fichier de suivi de l'execution, cree dans le meme dossier que le script.
$NomFichierLog = "Rapport-Utilisation-PowerBI.log"
$CheminLog = Join-Path $PSScriptRoot $NomFichierLog

# Mettre $true pour inclure les tentatives de connexion echouees.
$InclureLesEchecs = $false
# ============================================================================

# Identifiants officiels des services Microsoft :
# BI_AZURE_P2 = Power BI Pro
# BI_AZURE_P3 = Power BI Premium par utilisateur
$PowerBIProPlanId = [Guid]"70d33638-9c74-4d01-bfd3-562de28bd4ba"
$PowerBIPpuPlanId = [Guid]"0bf3c642-7bb5-4ccc-884e-59d09df0266c"
$PowerBIAppId      = "00000009-0000-0000-c000-000000000000"
$PowerBIDesktopApp = "Power BI Desktop"
$PowerBIDesktopAppAlternatif = "Microsoft Power BI Desktop"

function Write-Log {
    param(
        [string]$Message,
        [string]$Niveau = "INFO",
        [ConsoleColor]$Couleur = [ConsoleColor]::Gray
    )

    $Horodatage = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $Ligne = "[$Horodatage] [$Niveau] $Message"
    Write-Host $Ligne -ForegroundColor $Couleur
    Add-Content -Path $CheminLog -Value $Ligne -Encoding UTF8
}

# Le fichier de log est remis a zero a chaque execution.
Set-Content -Path $CheminLog -Value "" -Encoding UTF8
$Chronometre = [System.Diagnostics.Stopwatch]::StartNew()
Write-Log -Message "Demarrage du script $NomDuScript" -Couleur Cyan
Write-Log -Message "Rapport CSV prevu : $CheminCSV"
Write-Log -Message "Historique CSV prevu : $CheminHistoriqueCSV"

try {
    if ($TenantId -like "REMPLACER-*" -or $ClientId -like "REMPLACER-*") {
        throw "Completez TenantId et ClientId dans le bloc CONFIGURATION situe en haut du script."
    }

    if ($ModeAuthentification -notin @("Secret", "Certificat")) {
        throw "ModeAuthentification doit contenir Secret ou Certificat."
    }

    $ModulesRequis = @("Microsoft.Graph.Users", "Microsoft.Graph.Reports")
    foreach ($Module in $ModulesRequis) {
        Write-Log -Message "Verification du module $Module"
        if (-not (Get-Module -ListAvailable -Name $Module)) {
            Write-Log -Message "Installation du module $Module" -Couleur Yellow
            Install-Module $Module -Scope CurrentUser -Force -AllowClobber
        }
        Import-Module $Module
        Write-Log -Message "Module $Module charge" -Couleur Green
    }

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Log -Message "Connexion a Microsoft Graph avec application Azure" -Couleur Cyan

    if ($ModeAuthentification -eq "Certificat") {
        if ($CertificateThumbprint -like "REMPLACER-*" -or [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
            throw "Completez CertificateThumbprint dans le bloc CONFIGURATION."
        }

        Connect-MgGraph `
            -TenantId $TenantId `
            -ClientId $ClientId `
            -CertificateThumbprint $CertificateThumbprint `
            -NoWelcome
    }
    else {
        if ($ClientSecretEnClair -like "REMPLACER-*" -or [string]::IsNullOrWhiteSpace($ClientSecretEnClair)) {
            throw "Completez ClientSecretEnClair dans le bloc CONFIGURATION."
        }

        $ClientSecret = ConvertTo-SecureString $ClientSecretEnClair -AsPlainText -Force
        $IdentifiantsApplication = [PSCredential]::new($ClientId, $ClientSecret)
        Connect-MgGraph `
            -TenantId $TenantId `
            -ClientSecretCredential $IdentifiantsApplication `
            -NoWelcome
    }

    $Contexte = Get-MgContext
    if ($Contexte.AuthType -ne "AppOnly") {
        throw "La connexion Graph n’est pas de type AppOnly. Type obtenu : $($Contexte.AuthType)"
    }
    Write-Log -Message "Connexion Graph AppOnly reussie pour le tenant $TenantId" -Couleur Green

    Write-Log -Message "Chargement des utilisateurs du tenant" -Couleur Cyan

    $TousLesUtilisateurs = @(
        Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AccountEnabled,AssignedPlans
    )
    Write-Log -Message "$($TousLesUtilisateurs.Count) utilisateurs charges"

    Write-Log -Message "Filtrage des licences Power BI payantes" -Couleur Cyan

    $UtilisateursPowerBI = @(
        $TousLesUtilisateurs | Where-Object {
            $_.AssignedPlans | Where-Object {
                $_.CapabilityStatus -eq "Enabled" -and
                $_.ServicePlanId -in @($PowerBIProPlanId, $PowerBIPpuPlanId)
            }
        } | Sort-Object UserPrincipalName
    )

    if ($UtilisateursPowerBI.Count -eq 0) {
        Write-Log -Message "Aucun utilisateur avec une licence Power BI payante active" -Niveau "AVERTISSEMENT" -Couleur Yellow
        return
    }

    Write-Log -Message "$($UtilisateursPowerBI.Count) utilisateurs Power BI payants trouves" -Couleur Green

    $DateFinAnalyseUtc = (Get-Date).ToUniversalTime()
    $DateDebutAnalyseUtc = $DateFinAnalyseUtc.AddDays(-$NombreJoursAnalyse)
    $DateDebutFiltre = $DateDebutAnalyseUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
    $DateDebutAffichage = $DateDebutAnalyseUtc.ToLocalTime().ToString("dd/MM/yyyy HH:mm:ss")
    $DateFinAffichage = $DateFinAnalyseUtc.ToLocalTime().ToString("dd/MM/yyyy HH:mm:ss")

    Write-Log -Message "Periode demandee : du $DateDebutAffichage au $DateFinAffichage" -Couleur Cyan
    Write-Log -Message "Interrogation globale de Power BI Online" -Couleur Cyan

    $FiltreSucces = if ($InclureLesEchecs) { "" } else { " and status/errorCode eq 0" }
    $FiltreService = "appId eq '$PowerBIAppId' and createdDateTime ge $DateDebutFiltre$FiltreSucces"
    $FiltreDesktop = "(appDisplayName eq '$PowerBIDesktopApp' or appDisplayName eq '$PowerBIDesktopAppAlternatif') and createdDateTime ge $DateDebutFiltre$FiltreSucces"

    $EvenementsService = @(Get-MgAuditLogSignIn -Filter $FiltreService -All)
    Write-Log -Message "$($EvenementsService.Count) evenements Power BI Online recuperes" -Couleur Green

    Write-Log -Message "Interrogation globale de Power BI Desktop" -Couleur Cyan
    $EvenementsDesktop = @(Get-MgAuditLogSignIn -Filter $FiltreDesktop -All)
    Write-Log -Message "$($EvenementsDesktop.Count) evenements Power BI Desktop recuperes" -Couleur Green

    $TypeLicenceParUpn = @{}
    $UtilisateurParUpn = @{}
    $EvenementsParUpn = @{}

    foreach ($UtilisateurLicence in $UtilisateursPowerBI) {
        $CleUpn = $UtilisateurLicence.UserPrincipalName.ToLowerInvariant()
        $PlanPpu = $UtilisateurLicence.AssignedPlans | Where-Object {
            $_.CapabilityStatus -eq "Enabled" -and $_.ServicePlanId -eq $PowerBIPpuPlanId
        }
        $TypeLicenceUtilisateur = if ($PlanPpu) { "Power BI Premium par utilisateur" } else { "Power BI Pro" }
        $TypeLicenceParUpn[$CleUpn] = $TypeLicenceUtilisateur
        $UtilisateurParUpn[$CleUpn] = $UtilisateurLicence
        $EvenementsParUpn[$CleUpn] = New-Object System.Collections.ArrayList
    }

    $IdsVus = @{}
    $EvenementsPowerBI = @($EvenementsService) + @($EvenementsDesktop)
    foreach ($Evenement in $EvenementsPowerBI) {
        if ($null -eq $Evenement -or [string]::IsNullOrWhiteSpace($Evenement.UserPrincipalName)) {
            continue
        }

        $CleEvenement = $Evenement.UserPrincipalName.ToLowerInvariant()
        if (-not $UtilisateurParUpn.ContainsKey($CleEvenement)) {
            continue
        }

        $IdEvenement = "$($Evenement.Id)"
        if ($IdsVus.ContainsKey($IdEvenement)) {
            continue
        }

        $IdsVus[$IdEvenement] = $true
        [void]$EvenementsParUpn[$CleEvenement].Add($Evenement)
    }

    Write-Log -Message "$($IdsVus.Count) evenements concernent les utilisateurs Power BI payants" -Couleur Green

    $HistoriqueDetaille = foreach ($CleHistorique in $EvenementsParUpn.Keys) {
        $UtilisateurHistorique = $UtilisateurParUpn[$CleHistorique]
        $LicenceHistorique = $TypeLicenceParUpn[$CleHistorique]

        foreach ($EvenementHistorique in $EvenementsParUpn[$CleHistorique]) {
            $CanalHistorique = "Power BI Online"
            if ($EvenementHistorique.AppDisplayName -eq $PowerBIDesktopApp -or $EvenementHistorique.AppDisplayName -eq $PowerBIDesktopAppAlternatif) {
                $CanalHistorique = "Power BI Desktop"
            }

            $StatutHistorique = "Echec ($($EvenementHistorique.Status.ErrorCode))"
            if ($EvenementHistorique.Status.ErrorCode -eq 0) {
                $StatutHistorique = "Reussie"
            }

            $DateLocaleHistorique = ([DateTimeOffset]$EvenementHistorique.CreatedDateTime).ToLocalTime().ToString("dd/MM/yyyy HH:mm:ss zzz")
            $DateUtcHistorique = ([DateTimeOffset]$EvenementHistorique.CreatedDateTime).ToUniversalTime().ToString("o")
            $TypesHistorique = "$($EvenementHistorique.SignInEventTypes)"

            [PSCustomObject]@{
                DateHeureLocale     = $DateLocaleHistorique
                DateHeureUTC        = $DateUtcHistorique
                Nom                 = $UtilisateurHistorique.DisplayName
                Utilisateur         = $UtilisateurHistorique.UserPrincipalName
                LicencePowerBI      = $LicenceHistorique
                Canal               = $CanalHistorique
                Application         = $EvenementHistorique.AppDisplayName
                Interactif          = $EvenementHistorique.IsInteractive
                TypeConnexion       = $TypesHistorique
                Statut              = $StatutHistorique
                CodeErreur          = $EvenementHistorique.Status.ErrorCode
                Detail              = $EvenementHistorique.Status.FailureReason
                AdresseIP           = $EvenementHistorique.IpAddress
                Ressource           = $EvenementHistorique.ResourceDisplayName
                ResourceId          = $EvenementHistorique.ResourceId
                RequestId           = $EvenementHistorique.Id
                CorrelationId       = $EvenementHistorique.CorrelationId
                Systeme             = $EvenementHistorique.DeviceDetail.OperatingSystem
                Navigateur          = $EvenementHistorique.DeviceDetail.Browser
                Ville               = $EvenementHistorique.Location.City
                Pays                = $EvenementHistorique.Location.CountryOrRegion
                DateDebutExtraction = $DateDebutAffichage
                DateFinExtraction   = $DateFinAffichage
            }
        }
    }

    $HistoriqueDetaille = @($HistoriqueDetaille | Sort-Object DateHeureUTC -Descending)

    $Resultats = for ($Index = 0; $Index -lt $UtilisateursPowerBI.Count; $Index++) {
        $Utilisateur = $UtilisateursPowerBI[$Index]
        $Pourcentage = [math]::Round((($Index + 1) / $UtilisateursPowerBI.Count) * 100)
        $CleUpn = $Utilisateur.UserPrincipalName.ToLowerInvariant()
        $TypeLicence = $TypeLicenceParUpn[$CleUpn]
        $EvenementsUtilisateur = @($EvenementsParUpn[$CleUpn])

        Write-Progress -Activity "Creation du rapport Power BI" -Status "$($Index + 1)/$($UtilisateursPowerBI.Count) - $($Utilisateur.UserPrincipalName)" -PercentComplete $Pourcentage

        $ConnexionsServiceUtilisateur = @($EvenementsUtilisateur | Where-Object {
            $_.AppId -eq $PowerBIAppId
        })

        $ConnexionsDesktopUtilisateur = @($EvenementsUtilisateur | Where-Object {
            $_.AppDisplayName -eq $PowerBIDesktopApp -or $_.AppDisplayName -eq $PowerBIDesktopAppAlternatif
        })

        $ConnexionService = $ConnexionsServiceUtilisateur | Sort-Object CreatedDateTime -Descending | Select-Object -First 1
        $ConnexionDesktop = $ConnexionsDesktopUtilisateur | Sort-Object CreatedDateTime -Descending | Select-Object -First 1
        $PremiereConnexionService = $ConnexionsServiceUtilisateur | Sort-Object CreatedDateTime | Select-Object -First 1
        $PremiereConnexionDesktop = $ConnexionsDesktopUtilisateur | Sort-Object CreatedDateTime | Select-Object -First 1

        $NombreOnline = $ConnexionsServiceUtilisateur.Count
        $NombreDesktop = $ConnexionsDesktopUtilisateur.Count
        $NombreTotal = $NombreOnline + $NombreDesktop
        $OnlineInteractif = @($ConnexionsServiceUtilisateur | Where-Object { $_.IsInteractive -eq $true }).Count
        $OnlineNonInteractif = @($ConnexionsServiceUtilisateur | Where-Object { $_.IsInteractive -ne $true }).Count
        $DesktopInteractif = @($ConnexionsDesktopUtilisateur | Where-Object { $_.IsInteractive -eq $true }).Count
        $DesktopNonInteractif = @($ConnexionsDesktopUtilisateur | Where-Object { $_.IsInteractive -ne $true }).Count

        $JoursOnline = @($ConnexionsServiceUtilisateur | ForEach-Object {
            ([DateTimeOffset]$_.CreatedDateTime).ToLocalTime().ToString("yyyy-MM-dd")
        } | Sort-Object -Unique).Count

        $JoursDesktop = @($ConnexionsDesktopUtilisateur | ForEach-Object {
            ([DateTimeOffset]$_.CreatedDateTime).ToLocalTime().ToString("yyyy-MM-dd")
        } | Sort-Object -Unique).Count

        $ComparaisonCanaux = "Aucune utilisation detectee"
        if ($NombreOnline -gt 0 -and $NombreDesktop -gt 0) {
            $ComparaisonCanaux = "Desktop et Online"
        }
        elseif ($NombreOnline -gt 0) {
            $ComparaisonCanaux = "Online seulement"
        }
        elseif ($NombreDesktop -gt 0) {
            $ComparaisonCanaux = "Desktop seulement"
        }

        $CanalPrincipal = "Aucun"
        $PourcentageOnline = 0
        $PourcentageDesktop = 0
        if ($NombreTotal -gt 0) {
            $PourcentageOnline = [math]::Round(($NombreOnline / $NombreTotal) * 100, 1)
            $PourcentageDesktop = [math]::Round(($NombreDesktop / $NombreTotal) * 100, 1)

            if ($NombreOnline -gt $NombreDesktop) {
                $CanalPrincipal = "Power BI Online"
            }
            elseif ($NombreDesktop -gt $NombreOnline) {
                $CanalPrincipal = "Power BI Desktop"
            }
            else {
                $CanalPrincipal = "Egalite Desktop Online"
            }
        }

        $Connexion = @($ConnexionService, $ConnexionDesktop) | Where-Object {
            $null -ne $_
        } | Sort-Object CreatedDateTime -Descending | Select-Object -First 1

        $DateService = $null
        $DateDesktop = $null
        $PremiereDateService = $null
        $PremiereDateDesktop = $null
        $DateDerniereConnexion = $null
        $JoursDepuisConnexion = $null
        $IndiceUtilisation = "Non detectee sur la periode analysee"
        $DernierCanal = $null
        $TypeConnexion = $null
        $Interactive = $null
        $AdresseIP = $null
        $Ville = $null
        $Pays = $null
        $Systeme = $null
        $Navigateur = $null
        $Resultat = "Aucune connexion trouvee"
        $Detail = $null
        $CorrelationId = $null

        if ($ConnexionService) {
            $DateService = ([DateTimeOffset]$ConnexionService.CreatedDateTime).ToLocalTime().ToString("dd/MM/yyyy HH:mm:ss zzz")
        }
        if ($ConnexionDesktop) {
            $DateDesktop = ([DateTimeOffset]$ConnexionDesktop.CreatedDateTime).ToLocalTime().ToString("dd/MM/yyyy HH:mm:ss zzz")
        }
        if ($PremiereConnexionService) {
            $PremiereDateService = ([DateTimeOffset]$PremiereConnexionService.CreatedDateTime).ToLocalTime().ToString("dd/MM/yyyy HH:mm:ss zzz")
        }
        if ($PremiereConnexionDesktop) {
            $PremiereDateDesktop = ([DateTimeOffset]$PremiereConnexionDesktop.CreatedDateTime).ToLocalTime().ToString("dd/MM/yyyy HH:mm:ss zzz")
        }
        if ($Connexion) {
            $DateConnexionUtc = ([DateTimeOffset]$Connexion.CreatedDateTime).UtcDateTime
            $DateDerniereConnexion = ([DateTimeOffset]$Connexion.CreatedDateTime).ToLocalTime().ToString("dd/MM/yyyy HH:mm:ss zzz")
            $JoursDepuisConnexion = [math]::Floor(((Get-Date).ToUniversalTime() - $DateConnexionUtc).TotalDays)
            $IndiceUtilisation = "Oui - authentification detectee"
            $TypeConnexion = "$($Connexion.SignInEventTypes)"
            $Interactive = $Connexion.IsInteractive
            $AdresseIP = $Connexion.IpAddress
            $Ville = $Connexion.Location.City
            $Pays = $Connexion.Location.CountryOrRegion
            $Systeme = $Connexion.DeviceDetail.OperatingSystem
            $Navigateur = $Connexion.DeviceDetail.Browser
            $Detail = $Connexion.Status.FailureReason
            $CorrelationId = $Connexion.CorrelationId
            $Resultat = if ($Connexion.Status.ErrorCode -eq 0) { "Reussie" } else { "Echec ($($Connexion.Status.ErrorCode))" }
            $DernierCanal = if ($Connexion.AppId -eq $PowerBIAppId) { "Power BI Online" } else { "Power BI Desktop" }
        }

        Write-Log -Message "[$($Index + 1)/$($UtilisateursPowerBI.Count)] $($Utilisateur.UserPrincipalName) - Online: $NombreOnline - Desktop: $NombreDesktop - $ComparaisonCanaux" -Couleur Green

        [PSCustomObject]@{
            Nom                      = $Utilisateur.DisplayName
            Utilisateur              = $Utilisateur.UserPrincipalName
            CompteActif              = $Utilisateur.AccountEnabled
            LicencePowerBI           = $TypeLicence
            ComparaisonCanaux        = $ComparaisonCanaux
            CanalPrincipal           = $CanalPrincipal
            NombreConnexionsTotal    = $NombreTotal
            NombreConnexionsOnline   = $NombreOnline
            NombreConnexionsDesktop  = $NombreDesktop
            PourcentageOnline        = $PourcentageOnline
            PourcentageDesktop       = $PourcentageDesktop
            JoursActifsOnline        = $JoursOnline
            JoursActifsDesktop       = $JoursDesktop
            OnlineInteractif         = $OnlineInteractif
            OnlineNonInteractif      = $OnlineNonInteractif
            DesktopInteractif        = $DesktopInteractif
            DesktopNonInteractif     = $DesktopNonInteractif
            PremiereConnexionOnline  = $PremiereDateService
            DerniereConnexionService = $DateService
            PremiereConnexionDesktop = $PremiereDateDesktop
            DerniereConnexionDesktop = $DateDesktop
            DateDerniereConnexion    = $DateDerniereConnexion
            JoursDepuisConnexion     = $JoursDepuisConnexion
            IndiceUtilisationLicence = $IndiceUtilisation
            DernierCanalUtilise      = $DernierCanal
            TypeConnexion            = $TypeConnexion
            Interactive              = $Interactive
            AdresseIP                = $AdresseIP
            Ville                    = $Ville
            Pays                     = $Pays
            Systeme                  = $Systeme
            Navigateur               = $Navigateur
            Resultat                 = $Resultat
            Detail                   = $Detail
            CorrelationId            = $CorrelationId
            DateDebutExtraction      = $DateDebutAffichage
            DateFinExtraction        = $DateFinAffichage
        }
    }

    Write-Progress -Activity "Recherche des connexions Power BI" -Completed

    $DossierCSV = Split-Path -Parent $CheminCSV
    if ($DossierCSV -and -not (Test-Path $DossierCSV)) {
        New-Item -ItemType Directory -Path $DossierCSV -Force | Out-Null
    }

    $Resultats | Export-Csv -Path $CheminCSV -Delimiter ";" -NoTypeInformation -Encoding UTF8
    Write-Log -Message "Export CSV termine : $CheminCSV" -Couleur Green

    $HistoriqueDetaille | Export-Csv -Path $CheminHistoriqueCSV -Delimiter ";" -NoTypeInformation -Encoding UTF8
    Write-Log -Message "Export historique termine : $CheminHistoriqueCSV" -Couleur Green

    $Resultats | Format-Table Nom, Utilisateur, LicencePowerBI, ComparaisonCanaux, NombreConnexionsOnline, NombreConnexionsDesktop, CanalPrincipal -AutoSize

    Write-Host "`nRapport exporté : $CheminCSV" -ForegroundColor Green
    Write-Host "Historique exporte : $CheminHistoriqueCSV" -ForegroundColor Green
    Write-Host "Attention : ces evenements representent des authentifications, pas necessairement une utilisation active." -ForegroundColor Yellow
}
catch {
    Write-Log -Message "Arret du script : $($_.Exception.Message)" -Niveau "ERREUR" -Couleur Red
    Write-Error "Impossible de produire le rapport Power BI : $($_.Exception.Message)"
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    $Chronometre.Stop()
    Write-Log -Message "Fin du script - duree totale : $($Chronometre.Elapsed.ToString())" -Couleur Cyan
    Write-Log -Message "Fichier de log : $CheminLog"
}
