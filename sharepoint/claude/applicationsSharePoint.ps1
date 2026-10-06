#requires -Version 5.1

<#
.SYNOPSIS
    Reporte les applications SharePoint (solutions SPFx / apps du catalogue) installees sur le site source vers le site cible.

.DESCRIPTION
    1. Lit sur le site SOURCE les apps du catalogue tenant et du catalogue de collection de sites
       qui sont installees sur ce site.
    2. Verifie sur le site CIBLE que chaque app est disponible dans le catalogue (meme Id).
    3. Installe les apps manquantes sur le site cible, apres deploiement optionnel du package .sppkg.

    Tant que $Apply vaut $false, aucune modification n'est effectuee : seul le rapport est produit.

.IMPORTANT
    - Cross-tenant : une app doit d'abord etre deployee dans le catalogue de la cible. Placez les .sppkg
      dans $SppkgFolder (nom du fichier = titre ou Id de l'app) pour qu'ils soient ajoutes puis installes.
    - Les add-ins SharePoint classiques (modele "SharePoint Add-in") ne sont pas traites : modele obsolete.
    - Les donnees, listes, webparts de pages et parametres propres a une app ne sont pas migres ici.
    - Droits applicatifs : l'application doit pouvoir lire/installer des apps sur le site
      (FullControl sur le site) et, pour le catalogue tenant, acceder au catalogue d'applications.
    - Necessite PnP.PowerShell (PowerShell 7 recommande pour PnP 2.x/3.x).
#>

# ============================================================================
# CONFIGURATION A MODIFIER
# ============================================================================
$SourceSiteUrl = 'https://infoprodigital365.sharepoint.com/sites/fethygate'
$TargetSiteUrl = 'https://infoprodigital365.sharepoint.com/sites/fethiMercury'

$SourceTenantId = '2eea08b8-1972-447b-ad43-d044d042500a'
$SourceClientId = '821109a4-6e0e-48e8-b477-8f9b70aa32a4'
$SourceCertificatePath = 'C:\certs\MSGraphExchangeOnlineAuth20260717_2.pfx'
$SourceCertificatePassword = ''

# Laisser vide = memes valeurs que la source (copie dans le meme tenant).
$TargetTenantId = ''
$TargetClientId = ''
$TargetCertificatePath = ''
$TargetCertificatePassword = ''

# $false = simulation (rapport seulement). Mettre $true pour deployer/installer sur la cible.
$Apply = $false

# Catalogues a examiner : 'Tenant' et/ou 'Site' (catalogue de la collection de sites).
$Scopes = @('Tenant', 'Site')

# Dossier de packages .sppkg a deployer sur la cible quand l'app n'y est pas disponible (optionnel).
$SppkgFolder = ''

$baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$OutputDirectory = Join-Path $baseDir 'SharePoint-Apps'

# ============================================================================
# FIN DE LA CONFIGURATION
# ============================================================================
if ([string]::IsNullOrWhiteSpace($TargetTenantId)) { $TargetTenantId = $SourceTenantId }
if ([string]::IsNullOrWhiteSpace($TargetClientId)) { $TargetClientId = $SourceClientId }
if ([string]::IsNullOrWhiteSpace($TargetCertificatePath)) { $TargetCertificatePath = $SourceCertificatePath }
if ([string]::IsNullOrWhiteSpace($TargetCertificatePassword)) { $TargetCertificatePassword = $SourceCertificatePassword }

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
        throw 'Module PnP.PowerShell absent : Install-Module PnP.PowerShell -Scope CurrentUser'
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
    $script:ExitCode = 1
}
finally {
    if ($script:Results.Count -gt 0 -and (Test-Path -LiteralPath $OutputDirectory)) {
        $csv = Join-Path $OutputDirectory 'applications_resultats.csv'
        $script:Results | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
        Write-Log ("Rapport: {0} ({1} ligne(s))" -f $csv, $script:Results.Count)
    }
}

if ((Get-Variable -Name ExitCode -Scope Script -ErrorAction SilentlyContinue) -and $script:ExitCode -ne 0) { exit $script:ExitCode }
